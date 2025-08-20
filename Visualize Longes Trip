import os
import pandas as pd
import matplotlib.pyplot as plt

# Load files from the same directory as the script
script_dir = os.path.dirname(os.path.abspath(__file__))
layout_path = os.path.join(script_dir, 'Layout.xlsx')
pickdata_path = os.path.join(script_dir, 'PickData.xlsx')

# Load data
slot_df = pd.read_excel(layout_path)
pick_df = pd.read_excel(pickdata_path)

# Build item-to-slot and slot-to-coords lookup
slot_df_unique = slot_df.drop_duplicates(subset=['Current_Prime_Item'], keep='first')
item_to_slot = slot_df_unique.set_index('Current_Prime_Item')['Slot_ID'].to_dict()
slot_id_to_coords = slot_df_unique.set_index('Slot_ID')[['X', 'Y']].to_dict('index')
item_to_seq = slot_df_unique.set_index('Current_Prime_Item')['Pick_Seq'].to_dict()

# Filter out Full Pulls and non-Storage picks
pick_df = pick_df[(pick_df['Trip_Type'] != 'Full Pull') & (pick_df['Whse_Area'] == 'Storage')]

# Map Item to Slot_ID and Pick_Seq
pick_df['Slot_ID'] = pick_df['Item'].map(item_to_slot)
pick_df['Pick_Seq'] = pick_df['Item'].map(item_to_seq)

# Find the trip with the longest total distance
import math
def get_aisle_prefix(slot_id):
    return ''.join(c for c in str(slot_id) if not c.isdigit())
def calculate_achd_distance(slot1, slot2):
    if slot1 == slot2:
        return 0.0
    coords1 = slot_id_to_coords.get(slot1)
    coords2 = slot_id_to_coords.get(slot2)
    if not coords1 or not coords2:
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

trip_distances = {}
for trip, group in pick_df.groupby('Trip'):
    group = group.copy().sort_values('Pick_Seq')
    slots = group['Slot_ID'].tolist()
    total_distance = 0.0
    for i in range(1, len(slots)):
        slot1 = slots[i-1]
        slot2 = slots[i]
        if pd.notna(slot1) and pd.notna(slot2):
            total_distance += calculate_achd_distance(slot1, slot2)
    trip_distances[trip] = total_distance

# Find the trip with the max total distance
longest_trip = max(trip_distances, key=trip_distances.get)
longest_group = pick_df[pick_df['Trip'] == longest_trip].copy().sort_values('Pick_Seq')
route_slots = longest_group['Slot_ID'].tolist()
route_coords = [slot_id_to_coords.get(s) for s in route_slots if s in slot_id_to_coords]

# Plot warehouse layout
plt.figure(figsize=(12, 8))
plt.scatter(slot_df['X'], slot_df['Y'], c='lightgray', label='All Slots', s=10)

# Plot pick route using ACHD walking path
if route_coords and len(route_coords) > 1:
    for i in range(1, len(route_coords)):
        slot1 = route_slots[i-1]
        slot2 = route_slots[i]
        coords1 = slot_id_to_coords.get(slot1)
        coords2 = slot_id_to_coords.get(slot2)
        if not coords1 or not coords2:
            continue
        if get_aisle_prefix(slot1) == get_aisle_prefix(slot2):
            plt.plot([coords1['X'], coords2['X']], [coords1['Y'], coords2['Y']], c='blue', linewidth=2)
        else:
            x1, y1 = coords1['X'], coords1['Y']
            x2, y2 = coords2['X'], coords2['Y']
            crossover_y_values = [109, 1895, 3025, 3885]
            min_dist = float('inf')
            best_crossover = None
            for crossover_y in crossover_y_values:
                d = abs(y1 - crossover_y) + abs(x1 - x2) + abs(crossover_y - y2)
                if d < min_dist:
                    min_dist = d
                    best_crossover = crossover_y
            plt.plot([x1, x1], [y1, best_crossover], c='blue', linewidth=2)
            plt.plot([x1, x2], [best_crossover, best_crossover], c='blue', linewidth=2)
            plt.plot([x2, x2], [best_crossover, y2], c='blue', linewidth=2)
    # Dots at every stop/slot in the trip
    xs = [c['X'] for c in route_coords]
    ys = [c['Y'] for c in route_coords]
    plt.scatter(xs, ys, c='orange', s=40, label='Stops')
    # Start and end points
    plt.scatter(route_coords[0]['X'], route_coords[0]['Y'], c='green', s=80, label='Start')
    plt.scatter(route_coords[-1]['X'], route_coords[-1]['Y'], c='red', s=80, label='End')

plt.title(f'Longest Trip Route (Trip ID: {longest_trip})')
plt.xlabel('X (inches)')
plt.ylabel('Y (inches)')
plt.legend()
plt.tight_layout()

# Save and show
output_path = os.path.join(script_dir, 'longest_trip_route.png')
plt.savefig(output_path)
plt.show()
print(f"Saved route visualization to {output_path}") 
