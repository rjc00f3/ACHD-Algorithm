import pandas as pd
import math
import os
import re
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Font

def analyze_orders_achd(pick_data_path: str, slot_layout_path: str) -> pd.DataFrame:
    """
    Analyze orders using the ACHD algorithm.
    Args:
        pick_data_path (str): Path to the pick data Excel file (must have 'Order Number', 'Slot ID')
        slot_layout_path (str): Path to the slot layout Excel file (must have 'Slot_ID', 'Pick_Seq', 'X', 'Y')
    Returns:
        pd.DataFrame: Per-order distance metrics
    """
    print("Loading data files...")
    try:
        # Load data
        pick_df = pd.read_excel(pick_data_path)
        slot_df = pd.read_excel(slot_layout_path)
        print("\nPick data columns:", pick_df.columns.tolist())
        print("Slot layout columns:", slot_df.columns.tolist())
        # Ensure Current_Prime_Item is unique for mapping
        slot_df_unique = slot_df.drop_duplicates(subset=['Current_Prime_Item'], keep='first')
        item_to_slot = slot_df_unique.set_index('Current_Prime_Item')['Slot_ID'].to_dict()
        item_to_seq = slot_df_unique.set_index('Current_Prime_Item')['Pick_Seq'].to_dict()
        slot_id_to_coords = slot_df_unique.set_index('Slot_ID')[['X', 'Y']].to_dict('index')
        print(f"\nLoaded {len(item_to_slot)} item-to-slot mappings and {len(item_to_seq)} item-to-sequence mappings")
        def get_aisle_prefix(slot_id: str) -> str:
            return ''.join(c for c in slot_id if not c.isdigit())
        def calculate_achd_distance(slot1: str, slot2: str) -> float:
            if slot1 == slot2:
                return 0.0
            coords1 = slot_id_to_coords.get(slot1)
            coords2 = slot_id_to_coords.get(slot2)
            if not coords1 or not coords2:
                print(f"Warning: Missing coordinates for {slot1} or {slot2}")
                return 0.0
            if get_aisle_prefix(slot1) == get_aisle_prefix(slot2):
                return math.sqrt((coords1['X'] - coords2['X'])**2 + (coords1['Y'] - coords2['Y'])**2) / 12
            x1, y1 = coords1['X'], coords1['Y']
            x2, y2 = coords2['X'], coords2['Y']
            crossover_y_values = [109, 1895, 3025, 3885]
            min_dist = float('inf')
            for crossover_y in crossover_y_values:
                d = abs(y1 - crossover_y) + abs(x1 - x2) + abs(crossover_y - y2)
                min_dist = min(min_dist, d)
            return min_dist / 12
        print("\nAnalyzing orders...")
        # Filter out Full Pulls and non-Storage picks
        pick_df = pick_df[(pick_df['Trip_Type'] != 'Full Pull') & (pick_df['Whse_Area'] == 'Storage')]
        # Exclude trips with Pick_Slot PP09 or PP10
        if 'Pick_Slot' in pick_df.columns:
            trips_with_pp = pick_df[pick_df['Pick_Slot'].isin(['PP09', 'PP10'])]['Trip'].unique()
            pick_df = pick_df[~pick_df['Trip'].isin(trips_with_pp)]
        print(f"Filtered pick data: {len(pick_df)} rows remain after removing Full Pulls, non-Storage picks, and trips with PP09/PP10.")
        # Analyze by Trip_Category, excluding 'Full Pull trip'
        trip_categories = [cat for cat in pick_df['Trip_Category'].dropna().unique() if cat != 'Full Pull trip']
        output_dir = "warehouse_analysis_results"
        os.makedirs(output_dir, exist_ok=True)
        all_results = []
        for category in trip_categories:
            print(f"\nAnalyzing Trip Category: {category}")
            cat_df = pick_df[pick_df['Trip_Category'] == category].copy()
            # Map Item to Slot_ID, Pick_Seq, X, Y
            cat_df['Slot_ID'] = cat_df['Item'].map(item_to_slot)
            cat_df['Pick_Seq'] = cat_df['Item'].map(item_to_seq)
            cat_df['X'] = cat_df['Item'].map(lambda x: slot_id_to_coords.get(x, {}).get('X', float('nan')))
            cat_df['Y'] = cat_df['Item'].map(lambda x: slot_id_to_coords.get(x, {}).get('Y', float('nan')))
            results = []
            total_trips = len(cat_df['Trip'].unique())
            for i, (trip, group) in enumerate(cat_df.groupby('Trip'), 1):
                if i % 100 == 0:
                    print(f"Processing trip {i}/{total_trips}")
                group = group.copy()
                group = group.sort_values('Pick_Seq')
                slots = group['Slot_ID'].tolist()
                trip_distances = []
                for i in range(1, len(slots)):
                    slot1 = slots[i-1]
                    slot2 = slots[i]
                    d = 0.0
                    if pd.notna(slot1) and pd.notna(slot2):
                        d = calculate_achd_distance(slot1, slot2)
                    trip_distances.append(d)
                total_distance = sum(trip_distances)
                avg_distance = total_distance / max(len(trip_distances), 1)
                results.append({
                    'Trip_Category': category,
                    'Trip': trip,
                    'Total Distance': total_distance,
                    'Average Distance': avg_distance,
                    'Num Picks': len(slots)
                })
            result_df = pd.DataFrame(results)
            all_results.append(result_df)
            print(f"\nGlobal Statistics for {category}:")
            print(f"Total Trips Analyzed: {len(result_df)}")
            print(f"Average Distance per Trip: {result_df['Total Distance'].mean():.2f}")
            print(f"Median Distance per Trip: {result_df['Total Distance'].median():.2f}")
            print(f"90th Percentile Distance: {result_df['Total Distance'].quantile(0.9):.2f}")
        # Prepare to write all categories to a single Excel file
        combined_output_file = os.path.join(output_dir, "trip_analysis_results_ALL_CATEGORIES.xlsx")
        with pd.ExcelWriter(combined_output_file, engine='openpyxl') as writer:
            for result_df, category in zip(all_results, trip_categories):
                sheet_name = re.sub(r'[^\w\d ]', '', str(category))[:31]
                result_df.to_excel(writer, sheet_name=sheet_name, index=False)
            print(f"\nCombined results saved to: {combined_output_file}")
        return pd.concat(all_results, ignore_index=True)
    except Exception as e:
        print(f"Error during analysis: {str(e)}")
        raise

if __name__ == "__main__":
    script_dir = os.path.dirname(os.path.abspath(__file__))
    pick_data_path = os.path.join(script_dir, 'PickData.xlsx')
    slot_layout_path = os.path.join(script_dir, 'Layout.xlsx')
    analyze_orders_achd(pick_data_path, slot_layout_path) 