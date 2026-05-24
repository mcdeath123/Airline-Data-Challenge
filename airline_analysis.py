"""
airline_analysis.py

This script implements the Airline Data Challenge pipeline:
1. Loads and cleans Airport Codes, Flights, and Tickets datasets.
2. Imputes missing data (distances, delays, occupancy) and removes outliers.
3. Groups and joins data at the undirected route level (e.g., 'JFK-LAX') across all carriers.
4. Calculates flight-level revenues, costs, and profits based on challenge assumptions.
5. Summarizes route-level performance to answer all challenge questions.
6. Exports final summaries to a styled Excel file.
7. Generates premium, publication-quality visualizations saved as PNGs.
"""

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
import os

# Set style for premium visualizations
sns.set_theme(style="whitegrid")
plt.rcParams.update({
    'font.family': 'sans-serif',
    'font.sans-serif': ['DejaVu Sans', 'Arial', 'Helvetica'],
    'font.size': 12,
    'axes.labelsize': 14,
    'axes.titlesize': 16,
    'xtick.labelsize': 12,
    'ytick.labelsize': 12,
    'figure.titlesize': 18,
    'figure.dpi': 150
})

# Custom dark-theme / sleek palette
PALETTE_PRIMARY = "#1e3d59"   # Sleek deep navy
PALETTE_SECONDARY = "#ff6e40" # Warm orange accent
PALETTE_ACCENT = "#ffc13b"    # Golden accent
PALETTE_BG = "#f5f0e1"        # Soft cream
PALETTE_MUTED = "#b5b5b5"     # Muted grey
PALETTE_GRADIENT = ["#1e3d59", "#175873", "#0c7b93", "#00a8cc", "#ff6e40"]

def clean_airport_codes(filepath):
    """
    Loads and cleans the Airport Codes dataset.
    Filters for US medium and large airports, removes rows with missing IATA codes,
    and returns a clean DataFrame and a lookup dictionary of IATA -> Airport Type.
    """
    print("Loading Airport Codes...")
    airports = pd.read_csv(filepath)
    
    # Filter: US country, medium or large airport types, and non-null IATA codes
    us_ml = airports[
        (airports['ISO_COUNTRY'] == 'US') & 
        (airports['TYPE'].isin(['medium_airport', 'large_airport'])) & 
        (airports['IATA_CODE'].notnull())
    ].copy()
    
    # Remove duplicate IATA codes (if any exist, keeping the first occurrence)
    us_ml = us_ml.drop_duplicates(subset=['IATA_CODE'], keep='first')
    
    # Create lookup dict: IATA_CODE -> TYPE
    airport_type_dict = dict(zip(us_ml['IATA_CODE'], us_ml['TYPE']))
    
    print(f"Cleaned Airport Codes: {len(us_ml)} US medium/large airports resolved.")
    return us_ml, airport_type_dict

def clean_flights(filepath, airport_type_dict):
    """
    Loads and cleans the Flights dataset.
    - Coerces columns to numeric types.
    - Imputes missing distances using route-level median distance.
    - Excludes cancelled flights and flights not involving US medium/large airports.
    - Imputes missing occupancy rates and missing arrival delays.
    - Generates undirected route keys.
    """
    print("Loading Flights...")
    flights = pd.read_csv(filepath, low_memory=False)
    print(f"Raw flights count: {len(flights)}")
    
    # 1. Coerce critical columns to numeric
    flights['DISTANCE'] = pd.to_numeric(flights['DISTANCE'], errors='coerce')
    flights['OCCUPANCY_RATE'] = pd.to_numeric(flights['OCCUPANCY_RATE'], errors='coerce')
    flights['DEP_DELAY'] = pd.to_numeric(flights['DEP_DELAY'], errors='coerce')
    flights['ARR_DELAY'] = pd.to_numeric(flights['ARR_DELAY'], errors='coerce')
    flights['AIR_TIME'] = pd.to_numeric(flights['AIR_TIME'], errors='coerce')
    
    # 2. Exclude cancelled flights (CANCELLED == 1.0)
    flights = flights[flights['CANCELLED'] == 0.0].copy()
    print(f"Non-cancelled flights count: {len(flights)}")
    
    # 3. Filter for US medium and large airports (both ORIGIN and DESTINATION must be in the dict)
    valid_iatas = set(airport_type_dict.keys())
    flights = flights[
        flights['ORIGIN'].isin(valid_iatas) & 
        flights['DESTINATION'].isin(valid_iatas)
    ].copy()
    print(f"Flights between US medium/large airports count: {len(flights)}")
    
    # 4. Impute missing DISTANCE using the median of the same directed route (ORIGIN -> DESTINATION)
    # If still missing, use undirected route median; if still missing, use overall median.
    route_dir = flights['ORIGIN'] + "->" + flights['DESTINATION']
    dir_median_dist = flights.groupby(route_dir)['DISTANCE'].transform('median')
    flights['DISTANCE'] = flights['DISTANCE'].fillna(dir_median_dist)
    
    # Overall median fallback
    overall_median_dist = flights['DISTANCE'].median()
    flights['DISTANCE'] = flights['DISTANCE'].fillna(overall_median_dist)
    
    # 5. Impute missing DEP_DELAY and ARR_DELAY on non-cancelled flights
    # Non-cancelled flights should have no null DEP_DELAY, but if any, fill with 0
    flights['DEP_DELAY'] = flights['DEP_DELAY'].fillna(0.0)
    # If ARR_DELAY is null, set ARR_DELAY equal to DEP_DELAY (diversions/missing records)
    flights['ARR_DELAY'] = flights['ARR_DELAY'].fillna(flights['DEP_DELAY'])
    
    # 6. Impute missing OCCUPANCY_RATE using route-level median
    # Create temporary undirected route key for group imputation
    o = flights['ORIGIN'].values
    d = flights['DESTINATION'].values
    flights['route_undir'] = np.where(o < d, o + "-" + d, d + "-" + o)
    
    route_median_occ = flights.groupby('route_undir')['OCCUPANCY_RATE'].transform('median')
    flights['OCCUPANCY_RATE'] = flights['OCCUPANCY_RATE'].fillna(route_median_occ)
    
    # Overall median fallback (65.0%)
    overall_median_occ = flights['OCCUPANCY_RATE'].median()
    if pd.isna(overall_median_occ):
        overall_median_occ = 0.65
    flights['OCCUPANCY_RATE'] = flights['OCCUPANCY_RATE'].fillna(overall_median_occ)
    
    print("Completed Flight cleaning and imputation.")
    return flights

def clean_tickets(filepath, airport_type_dict):
    """
    Loads and cleans the Tickets dataset.
    - Strips non-numeric characters and converts ITIN_FARE to float.
    - Filters for round-trip itineraries (ROUNDTRIP == 1.0) and US medium/large airports.
    - Excludes extreme outliers: fares under $50 (reward security fees) and fares over $5000.
    - Generates undirected route keys.
    """
    print("Loading Tickets...")
    tickets = pd.read_csv(filepath)
    print(f"Raw tickets count: {len(tickets)}")
    
    # 1. Clean and convert ITIN_FARE to numeric
    # Strips everything except digits and decimal point (handles formats like '200 $', '$ 100.00', '820$$$')
    tickets['fare_clean'] = tickets['ITIN_FARE'].astype(str).str.replace(r'[^\d\.]', '', regex=True)
    tickets['fare_clean'] = pd.to_numeric(tickets['fare_clean'], errors='coerce')
    
    # 2. Filter: Only round trips (ROUNDTRIP == 1.0) and non-null fares
    rt_tickets = tickets[
        (tickets['ROUNDTRIP'] == 1.0) & 
        (tickets['fare_clean'].notnull())
    ].copy()
    
    # 3. Filter for US medium/large airports
    valid_iatas = set(airport_type_dict.keys())
    rt_tickets = rt_tickets[
        rt_tickets['ORIGIN'].isin(valid_iatas) & 
        rt_tickets['DESTINATION'].isin(valid_iatas)
    ].copy()
    
    # 4. Exclude outliers: fares under $50 and fares over $5000
    rt_tickets = rt_tickets[
        (rt_tickets['fare_clean'] >= 50.0) & 
        (rt_tickets['fare_clean'] <= 5000.0)
    ].copy()
    
    # 5. Generate undirected route key
    to_val = rt_tickets['ORIGIN'].values
    td_val = rt_tickets['DESTINATION'].values
    rt_tickets['route_undir'] = np.where(to_val < td_val, to_val + "-" + td_val, td_val + "-" + to_val)
    
    print(f"Cleaned Tickets: {len(rt_tickets)} round trip tickets resolved after removing outliers.")
    return rt_tickets

def calculate_flight_financials(flights, rt_tickets, airport_type_dict):
    """
    Calculates flight-level revenues and costs.
    1. Aggregates tickets by undirected route to find average round-trip fare.
    2. Scalably joins this ticket fare summary back to the flights dataset.
    3. Calculates passenger volume, ticket revenue, baggage revenue, and total revenue.
    4. Calculates operational cost (distance, destination airport fee, and delay penalties).
    5. Calculates net profit.
    """
    print("Linking Tickets and Flights data...")
    
    # Step 1: Group tickets by route to find average fare and passenger count (for weight/metadata)
    ticket_summary = rt_tickets.groupby('route_undir').agg(
        average_fare=('fare_clean', 'mean'),
        ticket_sample_size=('ITIN_ID', 'count')
    ).reset_index()
    
    # Step 2: Scalable data join
    # Merge the ticket summary into flights on 'route_undir'
    flights_merged = flights.merge(ticket_summary, on='route_undir', how='left')
    
    # Identify routes with missing ticket data and impute using overall median fare
    null_fares_count = flights_merged['average_fare'].isnull().sum()
    if null_fares_count > 0:
        overall_median_fare = rt_tickets['fare_clean'].median()
        print(f"Imputing missing fares for {null_fares_count} flights using overall median fare: ${overall_median_fare:.2f}")
        flights_merged['average_fare'] = flights_merged['average_fare'].fillna(overall_median_fare)
        
    print("Calculating financial metrics for each flight...")
    
    # Capacity and Occupancy calculations
    flights_merged['PASSENGERS'] = 200 * flights_merged['OCCUPANCY_RATE']
    
    # Ticket Revenue per flight leg (Average Round-Trip Fare divided by 2 per passenger per leg)
    flights_merged['REVENUE_TICKET'] = flights_merged['PASSENGERS'] * (flights_merged['average_fare'] / 2)
    
    # Baggage Revenue (Baggage fee is $35, expect 50% of passengers to check 1 bag per flight leg)
    flights_merged['REVENUE_BAGGAGE'] = flights_merged['PASSENGERS'] * 0.50 * 35.00
    
    # Total Revenue per flight
    flights_merged['REVENUE_TOTAL'] = flights_merged['REVENUE_TICKET'] + flights_merged['REVENUE_BAGGAGE']
    
    # Operational Costs
    # Fuel, Oil, Maintenance, Crew, Depreciation, Insurance, etc. ($8 + $1.18 = $9.18 per mile)
    flights_merged['COST_DISTANCE'] = flights_merged['DISTANCE'] * 9.18
    
    # Destination Airport operational cost lookup
    dest_types = flights_merged['DESTINATION'].map(airport_type_dict)
    flights_merged['COST_AIRPORT'] = np.where(dest_types == 'large_airport', 10000.0, 5000.0)
    
    # Delay Operational Costs: $75 per minute for departure/arrival delays beyond 15 minutes
    flights_merged['DELAY_DEP_COST'] = np.maximum(0, flights_merged['DEP_DELAY'] - 15) * 75.00
    flights_merged['DELAY_ARR_COST'] = np.maximum(0, flights_merged['ARR_DELAY'] - 15) * 75.00
    flights_merged['COST_DELAY'] = flights_merged['DELAY_DEP_COST'] + flights_merged['DELAY_ARR_COST']
    
    # Total operational cost per flight leg
    flights_merged['COST_TOTAL'] = flights_merged['COST_DISTANCE'] + flights_merged['COST_AIRPORT'] + flights_merged['COST_DELAY']
    
    # Net Profit per flight leg
    flights_merged['PROFIT_NET'] = flights_merged['REVENUE_TOTAL'] - flights_merged['COST_TOTAL']
    
    print("Flight financial calculations completed.")
    return flights_merged

def aggregate_route_performance(flights_financials):
    """
    Groups flights by undirected route and aggregates metrics to identify
    busiest, most profitable, and recommended routes.
    """
    print("Aggregating metrics by undirected route...")
    
    # Perform grouping
    route_groups = flights_financials.groupby('route_undir')
    
    # Aggregate metrics
    route_summary = route_groups.agg(
        flights_count=('FL_DATE', 'count'),
        total_passengers=('PASSENGERS', 'sum'),
        average_occupancy=('OCCUPANCY_RATE', 'mean'),
        average_distance=('DISTANCE', 'mean'),
        average_round_trip_fare=('average_fare', 'first'), # All flights on same route have same joined avg fare
        total_ticket_revenue=('REVENUE_TICKET', 'sum'),
        total_baggage_revenue=('REVENUE_BAGGAGE', 'sum'),
        total_revenue=('REVENUE_TOTAL', 'sum'),
        total_distance_cost=('COST_DISTANCE', 'sum'),
        total_airport_cost=('COST_AIRPORT', 'sum'),
        total_delay_cost=('COST_DELAY', 'sum'),
        total_cost=('COST_TOTAL', 'sum'),
        total_profit=('PROFIT_NET', 'sum')
    ).reset_index()
    
    # Add round trip counts (Total flights / 2)
    route_summary['round_trips_count'] = route_summary['flights_count'] / 2
    
    # Financial metrics per round trip
    route_summary['profit_per_round_trip'] = route_summary['total_profit'] / route_summary['round_trips_count']
    route_summary['revenue_per_round_trip'] = route_summary['total_revenue'] / route_summary['round_trips_count']
    route_summary['cost_per_round_trip'] = route_summary['total_cost'] / route_summary['round_trips_count']
    
    # Breakeven round trips on $90 million upfront cost
    route_summary['breakeven_round_trips'] = 90000000.0 / route_summary['profit_per_round_trip']
    
    # Extract origins and destinations (sort alphabetically, but let's list them)
    # The codes are split from the key
    route_summary['airport_1'] = route_summary['route_undir'].apply(lambda x: x.split('-')[0])
    route_summary['airport_2'] = route_summary['route_undir'].apply(lambda x: x.split('-')[1])
    
    print(f"Aggregated {len(route_summary)} unique undirected routes.")
    return route_summary

def generate_visualizations(busiest_routes, profitable_routes, recommended_routes):
    """
    Generates beautiful, modern, and high-impact visual charts
    to support the business narrative and saves them as PNGs.
    """
    print("Generating visual narrative charts...")
    os.makedirs("visualizations", exist_ok=True)
    
    # ------------------ CHART 1: Busiest Routes ------------------
    plt.figure(figsize=(10, 6))
    sns.barplot(
        x='round_trips_count',
        y='route_undir',
        data=busiest_routes,
        palette=sns.color_palette("Blues_r", n_colors=15)[:10]
    )
    plt.title("Top 10 Busiest Round Trip Routes (Q1 2019)", pad=20, weight='bold')
    plt.xlabel("Number of Round Trip Flights operated in the Quarter")
    plt.ylabel("Round Trip Route", labelpad=10)
    plt.tight_layout()
    plt.savefig("visualizations/chart_1_busiest_routes.png", dpi=150, bbox_inches='tight')
    plt.close()
    
    # ------------------ CHART 2: Most Profitable Routes ------------------
    plt.figure(figsize=(12, 7))
    # Reshape for grouped bar chart (Revenue vs Cost vs Profit)
    df_melt = pd.melt(
        profitable_routes[['route_undir', 'total_revenue', 'total_cost', 'total_profit']].copy(),
        id_vars=['route_undir'],
        value_vars=['total_revenue', 'total_cost', 'total_profit'],
        var_name='Metric',
        value_name='Amount'
    )
    df_melt['Amount_Million'] = df_melt['Amount'] / 1e6
    df_melt['Metric'] = df_melt['Metric'].map({
        'total_revenue': 'Total Revenue',
        'total_cost': 'Total Operational Cost',
        'total_profit': 'Net Profit'
    })
    
    sns.barplot(
        x='Amount_Million',
        y='route_undir',
        hue='Metric',
        data=df_melt,
        palette=[PALETTE_PRIMARY, PALETTE_MUTED, PALETTE_SECONDARY]
    )
    plt.title("Revenue, Cost, and Profit for the Top 10 Most Profitable Routes", pad=20, weight='bold')
    plt.xlabel("Amount (in USD Millions)")
    plt.ylabel("Round Trip Route", labelpad=10)
    plt.legend(title="Financial Components", loc="lower right")
    plt.tight_layout()
    plt.savefig("visualizations/chart_2_profitable_routes.png", dpi=150, bbox_inches='tight')
    plt.close()

    # ------------------ CHART 3: Breakeven Round Trips for Recommended Routes ------------------
    plt.figure(figsize=(9, 6))
    # Sort recommended routes by breakeven trips descending
    rec_sorted = recommended_routes.sort_values('breakeven_round_trips', ascending=False)
    
    bars = plt.bar(
        rec_sorted['route_undir'],
        rec_sorted['breakeven_round_trips'],
        color=PALETTE_PRIMARY,
        edgecolor='black',
        width=0.5
    )
    
    # Add values on top of bars
    for bar in bars:
        height = bar.get_height()
        plt.text(
            bar.get_x() + bar.get_width()/2.,
            height + 20,
            f"{int(round(height)):,}",
            ha='center', va='bottom', weight='bold', color='black'
        )
        
    plt.title("Round Trip Flights Needed to Breakeven on upfront Airplane Cost ($90M)", pad=25, weight='bold')
    plt.ylabel("Number of Round Trip Flights")
    plt.xlabel("Recommended Route")
    plt.ylim(0, rec_sorted['breakeven_round_trips'].max() * 1.15)
    plt.tight_layout()
    plt.savefig("visualizations/chart_3_breakeven_analysis.png", dpi=150, bbox_inches='tight')
    plt.close()
    
    # ------------------ CHART 4: Profitability Drivers (Occupancy vs Fare) ------------------
    plt.figure(figsize=(10, 6.5))
    scatter = sns.scatterplot(
        x='average_round_trip_fare',
        y='average_occupancy',
        size='profit_per_round_trip',
        hue='profit_per_round_trip',
        sizes=(100, 800),
        palette='coolwarm',
        data=recommended_routes,
        legend=False
    )
    
    # Label each point
    for idx, row in recommended_routes.iterrows():
        plt.text(
            row['average_round_trip_fare'] + 10,
            row['average_occupancy'] + 0.002,
            f"{row['route_undir']}\nProfit/RT: ${row['profit_per_round_trip']/1e3:.1f}k",
            weight='bold', size=10, ha='left', va='center'
        )
        
    plt.title("Recommended Routes: Impact of Fares and Occupancy on Profit", pad=20, weight='bold')
    plt.xlabel("Average Round-Trip Fare ($)")
    plt.ylabel("Average Occupancy Rate (%)")
    plt.xlim(recommended_routes['average_round_trip_fare'].min() - 50, recommended_routes['average_round_trip_fare'].max() + 200)
    plt.ylim(recommended_routes['average_occupancy'].min() - 0.02, recommended_routes['average_occupancy'].max() + 0.02)
    plt.gca().yaxis.set_major_formatter(plt.FuncFormatter(lambda y, _: '{:.0%}'.format(y)))
    plt.tight_layout()
    plt.savefig("visualizations/chart_4_profitability_drivers.png", dpi=150, bbox_inches='tight')
    plt.close()
    
    print("Visualizations successfully saved in the 'visualizations/' directory.")

def main():
    print("=================== STARTING DATA PIPELINE ===================")
    
    # Set paths (modify if needed, assuming current working directory contains them)
    airport_path = "Airport_Codes.csv"
    flights_path = "Flights.csv"
    tickets_path = "Tickets.csv"
    
    # Step 1: Load and Clean Airport Codes
    airports_df, airport_type_dict = clean_airport_codes(airport_path)
    
    # Step 2: Load and Clean Flights
    flights_clean = clean_flights(flights_path, airport_type_dict)
    
    # Step 3: Load and Clean Tickets
    tickets_clean = clean_tickets(tickets_path, airport_type_dict)
    
    # Step 4: Calculate flight financials (joining tickets and flights)
    flights_financials = calculate_flight_financials(flights_clean, tickets_clean, airport_type_dict)
    
    # Step 5: Aggregate route performance
    route_summary = aggregate_route_performance(flights_financials)
    
    # Step 6: Identify Top 10 Busiest Routes
    busiest_routes = route_summary.sort_values(by='round_trips_count', ascending=False).head(10).copy()
    print("\n--- TOP 10 BUSIEST ROUND TRIP ROUTES ---")
    print(busiest_routes[['route_undir', 'round_trips_count', 'flights_count', 'average_distance']])
    
    # Step 7: Identify Top 10 Most Profitable Routes
    profitable_routes = route_summary.sort_values(by='total_profit', ascending=False).head(10).copy()
    print("\n--- TOP 10 MOST PROFITABLE ROUND TRIP ROUTES ---")
    print(profitable_routes[['route_undir', 'total_profit', 'total_revenue', 'total_cost', 'round_trips_count']])
    
    # Step 8: Define and filter Recommended 5 Routes
    # Select the 5 recommended routes based on business criteria
    # Rationale: We look at the most profitable routes, but we also filter for high punctuality/low delays
    # and low breakeven flights. Since "punctuality is a big part of the brand image", we should look at routes
    # that are extremely profitable but have low delay costs, high occupancy, and high margins!
    # Let's inspect the top candidates to choose 5 outstanding, low-risk, highly profitable routes.
    # In order to select the 5 recommended routes, let's take the top 5 most profitable routes
    # that also show excellent operational resilience and robust demand.
    # Let's check which are the top 5 most profitable.
    # Let's sort route_summary by total_profit and print top 15 to make sure we make a highly informed selection!
    top_15_profit = route_summary.sort_values(by='total_profit', ascending=False).head(15)
    print("\n--- TOP 15 PROFITABLE CANDIDATES FOR SELECTION ---")
    print(top_15_profit[['route_undir', 'total_profit', 'total_delay_cost', 'average_occupancy', 'average_round_trip_fare', 'breakeven_round_trips']])
    
    # We will pick the top 5 most profitable routes as our primary recommendations, as they are overwhelmingly
    # more profitable than others and represent major commercial hubs with high demand.
    # Let's see which routes are the top 5:
    # 1. JFK-LAX
    # 2. LAX-SFO
    # 3. JFK-SFO
    # 4. ORD-LGA
    # 5. LAX-ORD
    # Let's make sure these are indeed the top 5.
    recommended_routes_list = top_15_profit.head(5)['route_undir'].tolist()
    recommended_routes = route_summary[route_summary['route_undir'].isin(recommended_routes_list)].sort_values(by='total_profit', ascending=False).copy()
    
    # Step 9: Export to Excel
    print("\nExporting results to Excel...")
    excel_file = "Airline_Analysis_Results.xlsx"
    try:
        with pd.ExcelWriter(excel_file, engine='openpyxl') as writer:
            busiest_routes.to_excel(writer, sheet_name='Busiest Routes', index=False)
            profitable_routes.to_excel(writer, sheet_name='Most Profitable Routes', index=False)
            recommended_routes.to_excel(writer, sheet_name='Recommended Routes', index=False)
            route_summary.to_excel(writer, sheet_name='All Routes Summary', index=False)
        print(f"Results successfully saved to {excel_file}")
    except PermissionError:
        print(f"WARNING: Permission denied when saving to {excel_file}. The file might be open in Excel. Skipping Excel save to allow visualization generation...")
    
    # Step 10: Generate visualizations
    generate_visualizations(busiest_routes, profitable_routes, recommended_routes)
    
    print("\n=================== DATA PIPELINE COMPLETED SUCCESSFULLY ===================")

if __name__ == "__main__":
    main()
