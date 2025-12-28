#!/usr/bin/env python3
"""
Script to create sample Excel files with data for the application.
This replaces the old JSON files with Excel format.
"""

import openpyxl
from openpyxl import Workbook
import json
import os

# Create data directory if it doesn't exist
os.makedirs('data', exist_ok=True)

# Sample Clients Data
clients_data = [
    {
        "id": 1,
        "name": "SFI",
        "destination": "SFI Depot",
        "itineraire": "Point A, Point B, Point C"
    },
    {
        "id": 2,
        "name": "Client B",
        "destination": "Warehouse B",
        "itineraire": "Point D, Point E, Point F"
    },
    {
        "id": 3,
        "name": "Client Test",
        "destination": "Test Location",
        "itineraire": "Start, Middle, End"
    },
    {
        "id": 4,
        "name": "ABC Company",
        "destination": "ABC Warehouse",
        "itineraire": "Route 1, Route 2"
    }
]

# Sample Drivers Data
drivers_data = [
    {
        "id": 1,
        "name": "Ahmed Benali",
        "cin": "AB123456",
        "phone": "0612345678",
        "vehicle": {
            "matricule": "12345-A-56",
            "model": "Mercedes Actros"
        }
    },
    {
        "id": 2,
        "name": "Mohamed Alami",
        "cin": "MA789012",
        "phone": "0623456789",
        "vehicle": {
            "matricule": "67890-B-12",
            "model": "Volvo FH"
        }
    },
    {
        "id": 3,
        "name": "Hassan Idrissi",
        "cin": "HI345678",
        "phone": "0634567890",
        "vehicle": {
            "matricule": "11111-C-34",
            "model": "Scania R450"
        }
    }
]

# Sample Convoyeurs Data
convoyeurs_data = [
    {
        "id": 1,
        "name": "Omar Tazi",
        "cin": "OT111222",
        "phone": "0645678901",
        "cce": "CCE001"
    },
    {
        "id": 2,
        "name": "Youssef Alaoui",
        "cin": "YA333444",
        "phone": "0656789012",
        "cce": "CCE002"
    },
    {
        "id": 3,
        "name": "Karim Bensaid",
        "cin": "KB555666",
        "phone": "0667890123"
    }
]

# Sample Products Data
products_data = [
    {
        "id": 1,
        "name": "Produit A",
        "unit": "Kg"
    },
    {
        "id": 2,
        "name": "Produit B",
        "unit": "Litre"
    },
    {
        "id": 3,
        "name": "Produit C",
        "unit": "Unité"
    },
    {
        "id": 4,
        "name": "Produit D",
        "unit": "Tonnes"
    },
    {
        "id": 5,
        "name": "Produit E",
        "unit": "Mètres"
    }
]

# History (empty for now)
history_data = []

def create_excel_file(filename, data, sheet_name):
    """Create an Excel file from data"""
    wb = Workbook()
    ws = wb.active
    ws.title = sheet_name
    
    if not data:
        # Create empty file with headers if data is empty
        if sheet_name == "Clients":
            ws.append(["id", "name", "destination", "itineraire"])
        elif sheet_name == "Conducteurs":
            ws.append(["id", "name", "cin", "phone", "vehicle.matricule", "vehicle.model"])
        elif sheet_name == "Convoyeurs":
            ws.append(["id", "name", "cin", "phone", "cce"])
        elif sheet_name == "Produits":
            ws.append(["id", "name", "unit"])
        elif sheet_name == "Historique":
            ws.append(["id", "documentNumber", "timestamp", "client", "driver", "products"])
    else:
        # Write headers
        if sheet_name == "Clients":
            ws.append(["id", "name", "destination", "itineraire"])
            for item in data:
                itineraire_str = item.get("itineraire", "")
                if isinstance(itineraire_str, list):
                    itineraire_str = ", ".join(itineraire_str)
                ws.append([
                    item.get("id", ""),
                    item.get("name", ""),
                    item.get("destination", ""),
                    itineraire_str
                ])
        elif sheet_name == "Conducteurs":
            ws.append(["id", "name", "cin", "phone", "vehicle.matricule", "vehicle.model"])
            for item in data:
                vehicle = item.get("vehicle", {})
                ws.append([
                    item.get("id", ""),
                    item.get("name", ""),
                    item.get("cin", ""),
                    item.get("phone", ""),
                    vehicle.get("matricule", ""),
                    vehicle.get("model", "")
                ])
        elif sheet_name == "Convoyeurs":
            ws.append(["id", "name", "cin", "phone", "cce"])
            for item in data:
                ws.append([
                    item.get("id", ""),
                    item.get("name", ""),
                    item.get("cin", ""),
                    item.get("phone", ""),
                    item.get("cce", "")
                ])
        elif sheet_name == "Produits":
            ws.append(["id", "name", "unit"])
            for item in data:
                ws.append([
                    item.get("id", ""),
                    item.get("name", ""),
                    item.get("unit", "")
                ])
        elif sheet_name == "Historique":
            ws.append(["id", "documentNumber", "timestamp", "client", "driver", "products"])
            for item in data:
                ws.append([
                    item.get("id", ""),
                    item.get("documentNumber", ""),
                    item.get("timestamp", ""),
                    json.dumps(item.get("client", {})),
                    json.dumps(item.get("driver", {})),
                    json.dumps(item.get("products", []))
                ])
    
    # Save file
    filepath = os.path.join("data", filename)
    wb.save(filepath)
    print(f"Created {filepath}")

# Create all Excel files
print("Creating Excel files with sample data...\n")

create_excel_file("clients.xlsx", clients_data, "Clients")
create_excel_file("drivers.xlsx", drivers_data, "Conducteurs")
create_excel_file("convoyeurs.xlsx", convoyeurs_data, "Convoyeurs")
create_excel_file("products.xlsx", products_data, "Produits")
create_excel_file("history.xlsx", history_data, "Historique")

print("\nAll Excel files created successfully!")
print("\nFiles created in data/ folder:")
print("  - clients.xlsx")
print("  - drivers.xlsx")
print("  - convoyeurs.xlsx")
print("  - products.xlsx")
print("  - history.xlsx")

