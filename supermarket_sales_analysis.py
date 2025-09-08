import pandas as pd
import sqlite3

# Step 1: Load dataset
df = pd.read_csv("supermarket_sales.csv")

# Step 2: Clean data
df = df.drop_duplicates()
df['Date'] = pd.to_datetime(df['Date'])

# Extract hour from Time
df['time_clean'] = df['Time'].astype(str).str.strip()
df['hour'] = pd.to_datetime(df['time_clean'], format="%H:%M", errors="coerce").dt.hour

# Step 3: Save into SQLite
conn = sqlite3.connect("supermarket_sales.db")
df.to_sql("sales", conn, if_exists="replace", index=False)

# Step 4: Queries
queries = {
    "Top_Product_Lines": """
        SELECT "Product line", SUM(Sales) AS total_sales
        FROM sales
        GROUP BY "Product line"
        ORDER BY total_sales DESC
        LIMIT 5;
    """,
    "Avg_Bill_Branch": """
        SELECT Branch, AVG(Sales) AS avg_bill
        FROM sales
        GROUP BY Branch
        ORDER BY avg_bill DESC;
    """,
    "Payment_Popularity": """
        SELECT Payment, COUNT(*) AS txn_count
        FROM sales
        GROUP BY Payment
        ORDER BY txn_count DESC;
    """,
    "Busiest_Hours": """
        SELECT hour, SUM(Sales) AS total_sales
        FROM sales
        GROUP BY hour
        ORDER BY total_sales DESC;
    """
}

# Step 5: Run queries & save to Excel
output_file = "Supermarket_Analysis.xlsx"
with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
    # Save each query to its own sheet
    results = {}
    for sheet_name, query in queries.items():
        result = pd.read_sql(query, conn)
        results[sheet_name] = result
        result.to_excel(writer, sheet_name=sheet_name, index=False)
        print(f"\n=== {sheet_name} ===")
        print(result)

    # Create a simple Dashboard sheet
    dashboard_data = {
        "Metric": [
            "Top Product Line",
            "Branch with Highest Avg Bill",
            "Most Popular Payment Method",
            "Busiest Hour"
        ],
        "Value": [
            results["Top_Product_Lines"].iloc[0]["Product line"],
            results["Avg_Bill_Branch"].iloc[0]["Branch"],
            results["Payment_Popularity"].iloc[0]["Payment"],
            results["Busiest_Hours"].iloc[0]["hour"]
        ]
    }
    dashboard_df = pd.DataFrame(dashboard_data)
    dashboard_df.to_excel(writer, sheet_name="Dashboard", index=False)

print(f"\n✅ Analysis completed and saved to {output_file}")
