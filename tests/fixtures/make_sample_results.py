"""Genera una results-DB SQLite de ejemplo con las 6 vistas que Arena crea.

Solo stdlib. Sirve para demostrar inspect_arena_results/read_arena_results
en Linux, donde las tools COM no pueden correr.
"""
import sqlite3
from pathlib import Path

VIEWS = {
    "ProjectQuery": "SELECT 'DemoProject' AS project, 'Model.doe' AS model, 3 AS replications",
    "OutputStatsQuery": "SELECT 'Tally 1' AS label, 12.5 AS average, 1.0 AS minimum, 30.0 AS maximum",
    "ContinuousTimeStatsQuery": "SELECT 'WIP' AS label, 4.2 AS average, 0.0 AS minimum, 9.0 AS maximum",
    "CounterStatsQuery": "SELECT 'PartsIn' AS label, 100 AS count_value",
    "DiscreteTimeStatsQuery": "SELECT 'QueueTime' AS label, 2.3 AS average",
    "FrequencyStatsQuery": "SELECT 'Resource' AS label, 'Busy' AS state, 0.75 AS frequency",
}

def main() -> None:
    out = Path(__file__).resolve().parent / "sample_results.db"
    if out.exists():
        out.unlink()
    conn = sqlite3.connect(out)
    conn.execute("CREATE TABLE replications (id INTEGER PRIMARY KEY, ended_ok INTEGER)")
    conn.executemany("INSERT INTO replications (ended_ok) VALUES (?)", [(1,), (1,), (1,)])
    for name, query in VIEWS.items():
        conn.execute(f"CREATE VIEW {name} AS {query}")
    conn.commit()
    conn.close()
    print(f"written {out} ({out.stat().st_size} bytes)")

if __name__ == "__main__":
    main()
