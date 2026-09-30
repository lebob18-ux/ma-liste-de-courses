#!/usr/bin/env python3
import os
import json
import requests
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from datetime import datetime
from dotenv import load_dotenv

# Load .env file
load_dotenv()

ACCESS_TOKEN = os.getenv("HELLOFRESH_ACCESS_TOKEN")
if not ACCESS_TOKEN:
    raise ValueError("❌ HELLOFRESH_ACCESS_TOKEN not found in .env file")

print(f"🔐 Using token for: banoar@gmail.com")
print(f"📍 Country: FR")

# API endpoints
BASE_URL = "https://www.hellofresh.fr"
RECIPES_ENDPOINT = f"{BASE_URL}/gw/api/recipes-meal-kit/saved"
PLANNED_ENDPOINT = f"{BASE_URL}/gw/api/recipes-meal-kit/planned"

headers = {
    "Authorization": f"Bearer {ACCESS_TOKEN}",
    "Accept": "application/json",
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
}

def fetch_recipes():
    """Fetch saved and planned recipes from HelloFresh"""
    all_recipes = []
    
    for endpoint_name, endpoint_url in [("Saved", RECIPES_ENDPOINT), ("Planned", PLANNED_ENDPOINT)]:
        try:
            print(f"\n📥 Fetching {endpoint_name} recipes...")
            response = requests.get(endpoint_url, headers=headers, timeout=10)
            response.raise_for_status()
            
            data = response.json()
            items = data.get("items", [])
            
            print(f"✅ Found {len(items)} {endpoint_name.lower()} recipes")
            
            for item in items:
                recipe = {
                    "id": item.get("id"),
                    "name": item.get("name"),
                    "slug": item.get("slug"),
                    "image_url": item.get("imagePath") or item.get("images", {}).get("LARGE"),
                    "servings": item.get("servings"),
                    "difficulty": item.get("difficulty"),
                    "time": item.get("time"),
                    "calories": item.get("calories"),
                    "description": item.get("description"),
                    "cuisineType": item.get("cuisineType"),
                    "allergens": ", ".join(item.get("allergens", [])),
                    "ingredients": [
                        {
                            "name": ing.get("name"),
                            "quantity": ing.get("quantity"),
                            "unit": ing.get("unit")
                        }
                        for ing in item.get("ingredients", [])
                    ]
                }
                all_recipes.append(recipe)
        
        except requests.exceptions.RequestException as e:
            print(f"⚠️  Error fetching {endpoint_name} recipes: {e}")
    
    return all_recipes

def create_excel_file(recipes):
    """Create an Excel file with recipes and ingredients"""
    
    wb = Workbook()
    if 'Sheet' in wb.sheetnames:
        wb.remove(wb['Sheet'])
    
    header_fill = PatternFill(start_color="2ECC71", end_color="2ECC71", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF", size=12)
    border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    left_alignment = Alignment(horizontal='left', vertical='top', wrap_text=True)
    center_alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    
    # Sheet 1: Summary
    ws_summary = wb.create_sheet("Recettes", 0)
    headers = ["Nom", "Difficulté", "Temps", "Servings", "Calories", "Type Cuisine", "Allergènes", "Lien Image"]
    for col_num, header in enumerate(headers, 1):
        cell = ws_summary.cell(row=1, column=col_num)
        cell.value = header
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = center_alignment
        cell.border = border
    
    for row_num, recipe in enumerate(recipes, 2):
        ws_summary.cell(row=row_num, column=1).value = recipe.get("name")
        ws_summary.cell(row=row_num, column=2).value = recipe.get("difficulty", "N/A")
        ws_summary.cell(row=row_num, column=3).value = recipe.get("time", "N/A")
        ws_summary.cell(row=row_num, column=4).value = recipe.get("servings", "N/A")
        ws_summary.cell(row=row_num, column=5).value = recipe.get("calories", "N/A")
        ws_summary.cell(row=row_num, column=6).value = recipe.get("cuisineType", "N/A")
        ws_summary.cell(row=row_num, column=7).value = recipe.get("allergens", "N/A")
        ws_summary.cell(row=row_num, column=8).value = recipe.get("image_url", "N/A")
        for col in range(1, len(headers) + 1):
            ws_summary.cell(row=row_num, column=col).border = border
            ws_summary.cell(row=row_num, column=col).alignment = left_alignment
    
    ws_summary.column_dimensions['A'].width = 30
    ws_summary.column_dimensions['B'].width = 12
    ws_summary.column_dimensions['C'].width = 12
    ws_summary.column_dimensions['D'].width = 10
    ws_summary.column_dimensions['E'].width = 10
    ws_summary.column_dimensions['F'].width = 15
    ws_summary.column_dimensions['G'].width = 30
    ws_summary.column_dimensions['H'].width = 50
    
    # Sheet 2: Detailed recipe data
    ws_detail = wb.create_sheet("Détails & Ingrédients", 1)
    current_row = 1
    
    for recipe in recipes:
        title_cell = ws_detail.cell(row=current_row, column=1)
        title_cell.value = recipe.get("name")
        title_cell.font = Font(bold=True, size=12, color="FFFFFF")
        title_cell.fill = PatternFill(start_color="3498DB", end_color="3498DB", fill_type="solid")
        ws_detail.merge_cells(f'A{current_row}:C{current_row}')
        current_row += 1

        info_data = [
            ("Difficulté:", recipe.get("difficulty", "N/A")),
            ("Temps:", recipe.get("time", "N/A")),
            ("Servings:", recipe.get("servings", "N/A")),
            ("Calories:", recipe.get("calories", "N/A")),
            ("Type Cuisine:", recipe.get("cuisineType", "N/A")),
            ("Allergènes:", recipe.get("allergens", "N/A")),
        ]
        for label, value in info_data:
            ws_detail.cell(row=current_row, column=1).value = label
            ws_detail.cell(row=current_row, column=1).font = Font(bold=True)
            ws_detail.cell(row=current_row, column=2).value = value
            current_row += 1

        current_row += 1
        ing_header_row = current_row
        ws_detail.cell(row=ing_header_row, column=1).value = "Ingrédient"
        ws_detail.cell(row=ing_header_row, column=2).value = "Quantité"
        ws_detail.cell(row=ing_header_row, column=3).value = "Unité"
        for col in range(1, 4):
            ws_detail.cell(row=ing_header_row, column=col).fill = PatternFill(start_color="E8F8F5", end_color="E8F8F5", fill_type="solid")
            ws_detail.cell(row=ing_header_row, column=col).font = Font(bold=True)
            ws_detail.cell(row=ing_header_row, column=col).border = border
        current_row += 1

        for ingredient in recipe.get("ingredients", []):
            ws_detail.cell(row=current_row, column=1).value = ingredient.get("name", "N/A")
            ws_detail.cell(row=current_row, column=2).value = ingredient.get("quantity", "N/A")
            ws_detail.cell(row=current_row, column=3).value = ingredient.get("unit", "N/A")
            for col in range(1, 4):
                ws_detail.cell(row=current_row, column=col).border = border
                ws_detail.cell(row=current_row, column=col).alignment = left_alignment
            current_row += 1

        current_row += 2
    
    ws_detail.column_dimensions['A'].width = 30
    ws_detail.column_dimensions['B'].width = 15
    ws_detail.column_dimensions['C'].width = 15

    # Sheet 3: Shopping list merged
    ws_shopping = wb.create_sheet("Liste de Courses", 2)
    ingredient_map = {}
    for recipe in recipes:
        for ing in recipe.get("ingredients", []):
            key = (ing.get("name") or "").lower().strip()
            if not key:
                continue
            if key not in ingredient_map:
                ingredient_map[key] = {
                    "name": ing.get("name"),
                    "quantity": ing.get("quantity"),
                    "unit": ing.get("unit"),
                    "count": 1
                }
            else:
                ingredient_map[key]["count"] += 1

    shopping_headers = ["Ingrédient", "Quantité Totale", "Unité", "Nombre de Recettes"]
    for col_num, header in enumerate(shopping_headers, 1):
        cell = ws_shopping.cell(row=1, column=col_num)
        cell.value = header
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = center_alignment
        cell.border = border

    for row_num, (_, ing_data) in enumerate(sorted(ingredient_map.items()), 2):
        ws_shopping.cell(row=row_num, column=1).value = ing_data["name"]
        ws_shopping.cell(row=row_num, column=2).value = ing_data["quantity"]
        ws_shopping.cell(row=row_num, column=3).value = ing_data["unit"]
        ws_shopping.cell(row=row_num, column=4).value = ing_data["count"]
        for col in range(1, len(shopping_headers) + 1):
            ws_shopping.cell(row=row_num, column=col).border = border
            ws_shopping.cell(row=row_num, column=col).alignment = left_alignment

    ws_shopping.column_dimensions['A'].width = 30
    ws_shopping.column_dimensions['B'].width = 15
    ws_shopping.column_dimensions['C'].width = 12
    ws_shopping.column_dimensions['D'].width = 20

    filename = f"HelloFresh_Recettes_{datetime.now().strftime('%Y-%m-%d_%H%M%S')}.xlsx"
    wb.save(filename)
    print(f"\n💾 Excel file created: {filename}")
    return filename

if __name__ == "__main__":
    try:
        recipes = fetch_recipes()
        if recipes:
            print(f"\n📊 Creating Excel file with {len(recipes)} recipes...")
            excel_file = create_excel_file(recipes)
            print(f"\n✨ Done! File created: {excel_file}")
        else:
            print("\n⚠️  No recipes found. Check your token or HelloFresh account.")
    except Exception as e:
        print(f"\n❌ Error: {e}")
        exit(1)
