import pyodbc
from pathlib import Path
from typing import List, Dict, Any
from src.utils.logging import get_logger
from src.utils.validation import validate_sql_safe

logger = get_logger("ps_search")

class PeopleSoftSearch:
    def __init__(self, connection_string: str, query_timeout: int = 30):
        self.conn_str = connection_string
        self.timeout = query_timeout

    def execute_query_from_file(self, sql_file: Path) -> List[Dict[str, Any]]:
        if not sql_file.exists():
            logger.error(f"SQL file not found: {sql_file}")
            return []
        
        with open(sql_file, 'r') as f:
            sql = f.read()
            
        return self.execute_query(sql)

    def execute_query(self, sql: str) -> List[Dict[str, Any]]:
        # Basic safety check
        if not validate_sql_safe(sql):
            logger.error("Write operations are not allowed. Query rejected.")
            raise ValueError("Safety violation: Query contains write operations.")

        results = []
        try:
            logger.info("Connecting to Database...")
            with pyodbc.connect(self.conn_str, timeout=self.timeout) as conn:
                cursor = conn.cursor()
                logger.info("Executing query...")
                cursor.execute(sql)
                
                if cursor.description:
                    columns = [column[0] for column in cursor.description]
                    for row in cursor.fetchall():
                        results.append(dict(zip(columns, row)))
                else:
                    # Query executed but returned no results/columns (e.g. set commands)
                    pass
                    
            logger.info(f"Query returned {len(results)} rows.")
        except pyodbc.Error as e:
            logger.error(f"Database error: {e}")
            # Re-raise or return empty depending on desired resilience
            # raise
        
        return results
