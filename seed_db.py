"""
SEED_DB.PY -- Initialize Camera Control database
=================================================
Creates tables and seeds test data for the camera control app.

Usage:
    DATABASE_URL=postgres://... py -3 seed_db.py

To add users or vehicles later, re-run with --add-user or --add-vehicle:
    DATABASE_URL=... py -3 seed_db.py --add-user "Name" PIN role location
    DATABASE_URL=... py -3 seed_db.py --add-vehicle ELD_ID VEH_NUM "Driver" location department
"""

import os
import sys
import psycopg2

DATABASE_URL = os.environ.get("DATABASE_URL", "")


def get_db():
    conn = psycopg2.connect(DATABASE_URL)
    conn.autocommit = True
    return conn


def create_tables(conn):
    """Create all required tables."""
    with conn.cursor() as cur:
        cur.execute("""
            CREATE TABLE IF NOT EXISTS users (
                id SERIAL PRIMARY KEY,
                name TEXT NOT NULL,
                pin TEXT UNIQUE NOT NULL,
                role TEXT NOT NULL CHECK (role IN ('admin', 'dispatch', 'driver')),
                location TEXT,
                created_at TIMESTAMP DEFAULT NOW()
            )
        """)

        cur.execute("""
            CREATE TABLE IF NOT EXISTS vehicles (
                id SERIAL PRIMARY KEY,
                eld_device_id TEXT NOT NULL,
                vehicle_number TEXT UNIQUE NOT NULL,
                driver_name TEXT NOT NULL,
                location TEXT NOT NULL,
                department TEXT NOT NULL DEFAULT 'Field',
                camera_state TEXT NOT NULL DEFAULT 'ON',
                last_toggled_at TIMESTAMP,
                last_toggled_by TEXT
            )
        """)

        cur.execute("""
            CREATE TABLE IF NOT EXISTS activity_log (
                id SERIAL PRIMARY KEY,
                timestamp TIMESTAMP NOT NULL DEFAULT NOW(),
                user_name TEXT NOT NULL,
                user_role TEXT NOT NULL,
                vehicle_number TEXT NOT NULL,
                action TEXT NOT NULL,
                motive_req_id TEXT,
                motive_status TEXT
            )
        """)

        # Create indexes for common queries
        cur.execute("CREATE INDEX IF NOT EXISTS idx_log_timestamp ON activity_log (timestamp DESC)")
        cur.execute("CREATE INDEX IF NOT EXISTS idx_log_user ON activity_log (user_name)")
        cur.execute("CREATE INDEX IF NOT EXISTS idx_log_vehicle ON activity_log (vehicle_number)")
        cur.execute("CREATE INDEX IF NOT EXISTS idx_vehicles_location ON vehicles (location)")
        cur.execute("CREATE INDEX IF NOT EXISTS idx_vehicles_driver ON vehicles (driver_name)")

    print("Tables created successfully.")


def seed_users(conn):
    """Seed test users."""
    users = [
        ("Admin", "1000", "admin", None),
        ("Dispatch Midland", "2001", "dispatch", "Midland"),
        ("Dispatch Laredo", "2002", "dispatch", "Laredo"),
        ("Dispatch Bryan", "2003", "dispatch", "Bryan"),
        ("Dispatch Hobbs", "2004", "dispatch", "Hobbs"),
        ("Dispatch Kilgore", "2005", "dispatch", "Kilgore"),
        ("Dispatch Jourdanton", "2006", "dispatch", "Jourdanton"),
        ("Bob Stokes", "3001", "driver", "Midland"),
        ("Jose Gonzalez", "3002", "driver", "Laredo"),
    ]

    with conn.cursor() as cur:
        for name, pin, role, location in users:
            cur.execute(
                "INSERT INTO users (name, pin, role, location) VALUES (%s, %s, %s, %s) "
                "ON CONFLICT (pin) DO UPDATE SET name = EXCLUDED.name, role = EXCLUDED.role, location = EXCLUDED.location",
                (name, pin, role, location))
            print(f"  User: {name} (PIN: {pin}, role: {role}, location: {location or 'all'})")

    print(f"Seeded {len(users)} users.")


def seed_vehicles(conn):
    """Seed test vehicles."""
    vehicles = [
        ("1674558", "2294C", "Bob Stokes", "Midland", "Field"),
        ("1680021", "2135C", "Jose Gonzalez", "Laredo", "Field"),
    ]

    with conn.cursor() as cur:
        for eld_id, vnum, driver, location, dept in vehicles:
            cur.execute(
                "INSERT INTO vehicles (eld_device_id, vehicle_number, driver_name, location, department) "
                "VALUES (%s, %s, %s, %s, %s) "
                "ON CONFLICT (vehicle_number) DO UPDATE SET eld_device_id = EXCLUDED.eld_device_id, "
                "driver_name = EXCLUDED.driver_name, location = EXCLUDED.location, department = EXCLUDED.department",
                (eld_id, vnum, driver, location, dept))
            print(f"  Vehicle: {vnum} ({driver}, {location})")

    print(f"Seeded {len(vehicles)} vehicles.")


def add_user(conn, name, pin, role, location):
    """Add a single user."""
    with conn.cursor() as cur:
        cur.execute(
            "INSERT INTO users (name, pin, role, location) VALUES (%s, %s, %s, %s) "
            "ON CONFLICT (pin) DO UPDATE SET name = EXCLUDED.name, role = EXCLUDED.role, location = EXCLUDED.location",
            (name, pin, role, location if location != "none" else None))
    print(f"Added user: {name} (PIN: {pin}, role: {role}, location: {location})")


def add_vehicle(conn, eld_id, vnum, driver, location, department):
    """Add a single vehicle."""
    with conn.cursor() as cur:
        cur.execute(
            "INSERT INTO vehicles (eld_device_id, vehicle_number, driver_name, location, department) "
            "VALUES (%s, %s, %s, %s, %s) "
            "ON CONFLICT (vehicle_number) DO UPDATE SET eld_device_id = EXCLUDED.eld_device_id, "
            "driver_name = EXCLUDED.driver_name, location = EXCLUDED.location, department = EXCLUDED.department",
            (eld_id, vnum, driver, location, department))
    print(f"Added vehicle: {vnum} ({driver}, {location})")


def show_status(conn):
    """Show current DB contents."""
    with conn.cursor() as cur:
        cur.execute("SELECT COUNT(*) FROM users")
        print(f"\nUsers: {cur.fetchone()[0]}")
        cur.execute("SELECT name, pin, role, location FROM users ORDER BY role, name")
        for row in cur.fetchall():
            print(f"  {row[0]:<25} PIN: {row[1]:<6} {row[2]:<10} {row[3] or 'all'}")

        cur.execute("SELECT COUNT(*) FROM vehicles")
        print(f"\nVehicles: {cur.fetchone()[0]}")
        cur.execute("SELECT vehicle_number, driver_name, location, camera_state FROM vehicles ORDER BY location, vehicle_number")
        for row in cur.fetchall():
            print(f"  {row[0]:<12} {row[1]:<25} {row[2]:<15} Camera: {row[3]}")

        cur.execute("SELECT COUNT(*) FROM activity_log")
        print(f"\nLog entries: {cur.fetchone()[0]}")


def main():
    if not DATABASE_URL:
        print("ERROR: DATABASE_URL environment variable not set.")
        print("Usage: DATABASE_URL=postgres://... py -3 seed_db.py")
        sys.exit(1)

    conn = get_db()

    if len(sys.argv) > 1 and sys.argv[1] == "--add-user":
        if len(sys.argv) < 6:
            print("Usage: seed_db.py --add-user \"Name\" PIN role location")
            print("  role: admin, dispatch, or driver")
            print("  location: yard name or 'none' for admin")
            sys.exit(1)
        add_user(conn, sys.argv[2], sys.argv[3], sys.argv[4], sys.argv[5])

    elif len(sys.argv) > 1 and sys.argv[1] == "--add-vehicle":
        if len(sys.argv) < 7:
            print("Usage: seed_db.py --add-vehicle ELD_ID VEH_NUM \"Driver Name\" location department")
            sys.exit(1)
        add_vehicle(conn, sys.argv[2], sys.argv[3], sys.argv[4], sys.argv[5], sys.argv[6])

    elif len(sys.argv) > 1 and sys.argv[1] == "--status":
        show_status(conn)

    else:
        print("Creating tables...")
        create_tables(conn)
        print("\nSeeding users...")
        seed_users(conn)
        print("\nSeeding vehicles...")
        seed_vehicles(conn)

    print("\n--- Current Status ---")
    show_status(conn)
    conn.close()


if __name__ == "__main__":
    main()
