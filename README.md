# Warehouse Distance & Route Analysis (Aisle Constrained Hybrid Distance Algorithm)

This repository contains Python tools for analyzing warehouse order travel distances and visualizing picking routes.  
The goal is to measure efficiency, identify long-travel orders, and support slotting optimization using real-world warehouse layouts.

---

## 📂 Contents
- **OrderDistanceCalculator.py**  
  Analyzes warehouse pick trips using a custom **ACHD distance algorithm**.  
  Produces per-trip metrics (total distance, average distance, # of picks) and exports results to Excel.  

- **VisualizeLongestTripRoute.py**  
  Plots the longest trip found in the data, showing actual walking paths across aisles with crossover points.  
  Saves a `.png` image of the warehouse map with the route highlighted.  

---

## 🚀 Features
- Cleans raw pick data (removes Full Pulls, non-Storage trips, and PP09/PP10 problem slots).  
- Maps **items → slots → coordinates**.  
- Custom **ACHD distance calculation** with aisle crossover logic.  
- Per-category Excel reports with distance statistics.  
- Visualization of the **longest trip** with start/end markers and route path.  
- Practical tool for **slotting optimization** and reducing labor travel distance.  

---

## 🛠 Requirements
- Python 3.9+  
- pandas  
- openpyxl  
- matplotlib  
