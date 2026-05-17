"""
Oracle Health Check Script
Monitors sessions, deadlocks, and tablespace usage
"""

import cx_Oracle

def health_check(connection):
    """Run comprehensive health checks"""
    
    checks = {
        'sessions': """
            SELECT COUNT(*) as active_sessions 
            FROM v$session 
            WHERE status = 'ACTIVE' AND type = 'USER'
        """,
        'deadlocks': """
            SELECT COUNT(*) as deadlock_count 
            FROM v$session 
            WHERE blocking_session IS NOT NULL
        """,
        'critical_tablespaces': """
            SELECT tablespace_name, 
                   ROUND((used_space/tablespace_size)*100, 2) as pct_used
            FROM dba_tablespace_usage_metrics
            WHERE (used_space/tablespace_size)*100 > 1
        """
    }
    
    cursor = connection.cursor()
    
    print("=" * 50)
    print("ORACLE DATABASE HEALTH CHECK")
    print("=" * 50)
    
    # Check 1: Active Sessions
    print("\n1. ACTIVE SESSIONS:")
    cursor.execute(checks['sessions'])
    sessions = cursor.fetchone()[0]
    status = "OK" if sessions < 100 else "WARNING" if sessions < 200 else "CRITICAL"
    print(f"   Active Sessions: {sessions} [{status}]")
    
    # Check 2: Deadlocks
    print("\n2. DEADLOCKS:")
    cursor.execute(checks['deadlocks'])
    deadlocks = cursor.fetchone()[0]
    status = "OK" if deadlocks == 0 else "CRITICAL"
    print(f"   Blocking Sessions: {deadlocks} [{status}]")
    
    # Check 3: Tablespace Usage
    print("\n3. TABLESPACE USAGE (>1%):")
    cursor.execute(checks['critical_tablespaces'])
    critical_tbs = cursor.fetchall()
    
    if critical_tbs:
        for tbs_name, pct_used in critical_tbs:
            print(f"   ⚠ {tbs_name}: {pct_used}% [CRITICAL]")
    else:
        print("   All tablespaces healthy [OK]")
    
    print("\n" + "=" * 50)
    cursor.close()

if __name__ == "__main__":
    # Update with your credentials
    conn = cx_Oracle.connect("hr/oracle123@localhost:1521/orcl19b")
    health_check(conn)
    conn.close()