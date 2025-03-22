import json
import time
import os
import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd
import numpy as np
import requests
from datetime import datetime
from collections import deque
from threading import Lock
import xlsxwriter

# Global strategy parameters
ATM_ADDER = 150
INCREMENTAL_ADDER = 50
NUM_BASES = 8
OFFSET_VALUES = [250, 300, 350]
RATIO_SETS = [[1, 5], [1, 4], [2, 7], [1, 3], [2, 5], [3, 7], [1, 2], [4, 7], [2, 3]]

# API and symbol configuration
API_TOKENS = [
    "eyJ0eXAiOiJKV1QiLCJrZXlfaWQiOiJza192MS4wIiwiYWxnIjoiSFMyNTYifQ.eyJzdWIiOiJBSjQ5NzIiLCJqdGkiOiI2N2RlMjgwODJkY2M5ZjMzMzc2YTkyMzEiLCJpc011bHRpQ2xpZW50IjpmYWxzZSwiaWF0IjoxNzQyNjEyNDg4LCJpc3MiOiJ1ZGFwaS1nYXRld2F5LXNlcnZpY2UiLCJleHAiOjE3NDI2ODA4MDB9.TMywvV5JO8jR3AYhjI_tvs0R76lIN38yTjIc64MHRdo",
    "eyJ0eXAiOiJKV1QiLCJrZXlfaWQiOiJza192MS4wIiwiYWxnIjoiSFMyNTYifQ.eyJzdWIiOiIzSkNZTUciLCJqdGkiOiI2N2RlMmE5YjNkZGQ3YTZhM2Q0MGIzZmMiLCJpc011bHRpQ2xpZW50IjpmYWxzZSwiaWF0IjoxNzQyNjEzMTQ3LCJpc3MiOiJ1ZGFwaS1nYXRld2F5LXNlcnZpY2UiLCJleHAiOjE3NDI2ODA4MDB9.qVBjOEoTiBvVyZz077PurrgN9pWpjXQqvui0WDFnkpM"
]

# List of instruments to analyze
INSTRUMENTS = [
    {
        "instrument_key": "NSE_INDEX|Nifty 50",
        "symbol_name": "NIFTY",
        "lot_size": 75,
        "expiry_date": "2025-03-27"
    }
    # Add more instruments as needed
]

# Global variables
global_results = []  # Stores analyzed option data
status_text = None  # Reference to status label

# Token rotation management
token_rotation = {
    'tokens': deque(API_TOKENS),
    'current_token_index': 0,
    'token_locks': {token: Lock() for token in API_TOKENS},
    'token_api_calls': {token: {
        'second': {'count': 0, 'reset_time': datetime.now()},
        'minute': {'count': 0, 'reset_time': datetime.now()},
        '30min': {'count': 0, 'reset_time': datetime.now()}
    } for token in API_TOKENS},
    'master_lock': Lock()
}

# Get next available token
def get_next_token():
    """Get the next available token with token rotation"""
    global token_rotation

    with token_rotation['master_lock']:
        # Get the next token and rotate
        token = token_rotation['tokens'][0]
        token_rotation['tokens'].rotate(-1)
        return token

# Rate limiting functions
def wait_for_rate_limit(token):
    """Wait if necessary to respect API rate limits"""
    global token_rotation

    with token_rotation['token_locks'][token]:
        api_calls = token_rotation['token_api_calls'][token]
        now = datetime.now()

        # Reset counters if time windows have passed
        if (now - api_calls['second']['reset_time']).total_seconds() >= 1:
            api_calls['second'] = {'count': 0, 'reset_time': now}

        if (now - api_calls['minute']['reset_time']).total_seconds() >= 60:
            api_calls['minute'] = {'count': 0, 'reset_time': now}

        if (now - api_calls['30min']['reset_time']).total_seconds() >= 1800:
            api_calls['30min'] = {'count': 0, 'reset_time': now}

        # Check if we're approaching limits and wait if necessary
        if api_calls['second']['count'] >= 45:  # Buffer of 5 requests
            wait_time = 1.0 - (now - api_calls['second']['reset_time']).total_seconds()
            if wait_time > 0:
                print(f"Rate limit approaching for token (per-second), waiting {wait_time:.2f}s")
                time.sleep(wait_time)
                # Reset after waiting
                api_calls['second'] = {'count': 0, 'reset_time': datetime.now()}

        if api_calls['minute']['count'] >= 450:  # Buffer of 50 requests
            wait_time = 60.0 - (now - api_calls['minute']['reset_time']).total_seconds()
            if wait_time > 0:
                print(f"Rate limit approaching for token (per-minute), waiting {wait_time:.2f}s")
                time.sleep(wait_time)
                # Reset after waiting
                api_calls['minute'] = {'count': 0, 'reset_time': datetime.now()}
                api_calls['second'] = {'count': 0, 'reset_time': datetime.now()}

        if api_calls['30min']['count'] >= 1900:  # Buffer of 100 requests
            wait_time = 1800.0 - (now - api_calls['30min']['reset_time']).total_seconds()
            if wait_time > 0:
                print(f"Rate limit approaching for token (30-minute window), waiting {wait_time:.2f}s")
                time.sleep(wait_time)
                # Reset after waiting
                api_calls['30min'] = {'count': 0, 'reset_time': datetime.now()}
                api_calls['minute'] = {'count': 0, 'reset_time': datetime.now()}
                api_calls['second'] = {'count': 0, 'reset_time': datetime.now()}

def update_rate_counters(token):
    """Update API call counters after making a request"""
    global token_rotation

    with token_rotation['token_locks'][token]:
        api_calls = token_rotation['token_api_calls'][token]
        api_calls['second']['count'] += 1
        api_calls['minute']['count'] += 1
        api_calls['30min']['count'] += 1

def fetch_option_chain(instrument_key, expiry_date):
    """
    Fetch option chain data from Upstox API with token rotation

    Parameters:
    -----------
    instrument_key : str
        Instrument key
    expiry_date : str
        Expiry date (format: 2025-03-27)

    Returns:
    --------
    dict
        Option chain data
    """
    url = 'https://api.upstox.com/v2/option/chain'

    # Get a token from the rotation system
    token = get_next_token()

    headers = {'Accept': 'application/json', 'Authorization': f'Bearer {token}'}
    params = {'instrument_key': instrument_key, 'expiry_date': expiry_date}

    print(f"Making request to {url}")
    print(f"Parameters: {params}")

    # Apply rate limiting before API call
    wait_for_rate_limit(token)

    try:
        response = requests.get(url, params=params, headers=headers, timeout=10)
        # Update rate limit counters
        update_rate_counters(token)

        if response.status_code == 200:
            print(f"Successfully fetched data for {instrument_key} {expiry_date}")
            return response.json()
        else:
            print(f"Error fetching data: {response.status_code}")
            print(f"Response text: {response.text}")

            # If first token failed, try with a different token
            different_token = get_next_token()

            # Make sure we got a different token
            if different_token != token:
                print("Retrying with different token...")
                headers = {'Accept': 'application/json', 'Authorization': f'Bearer {different_token}'}

                # Apply rate limiting for the second token
                wait_for_rate_limit(different_token)

                response = requests.get(url, params=params, headers=headers, timeout=10)
                # Update rate limit counters
                update_rate_counters(different_token)

                if response.status_code == 200:
                    print(f"Retry successful for {instrument_key} {expiry_date}")
                    return response.json()
                else:
                    print(f"Retry also failed with status: {response.status_code}")
                    print(f"Retry response text: {response.text}")

            return None

    except Exception as e:
        print(f"Exception in fetch_option_chain: {str(e)}")
        return None

def analyze_option_chain(option_chain_data, symbol_name="NIFTY", lot_size=50):
    """
    Analyze call options from option chain data

    Parameters:
    -----------
    option_chain_data : dict
        Option chain data containing calls
    symbol_name : str
        Name of the symbol
    lot_size : int
        Lot size for the symbol

    Returns:
    --------
    dict
        Results of the analysis
    """
    try:
        global global_results

        # Check if data is in Upstox API format
        if isinstance(option_chain_data, dict) and "data" in option_chain_data:
            # Extract underlying price
            underlying_price = option_chain_data["data"][0]["underlying_spot_price"]

            # Extract call options data
            calls_data = []
            for item in option_chain_data["data"]:
                if "call_options" in item and "market_data" in item["call_options"]:
                    market_data = item["call_options"]["market_data"]
                    # Skip if bid or ask is not available
                    if market_data["bid_price"] <= 0 or market_data["ask_price"] <= 0:
                        continue

                    calls_data.append({
                        "strike": item["strike_price"],
                        "bidPrice": market_data["bid_price"],
                        "askPrice": market_data["ask_price"],
                        "instrument_key": item["call_options"]["instrument_key"]
                    })
        else:
            # Use sample data format
            calls_data = option_chain_data.get('calls', [])
            underlying_price = option_chain_data.get('underlyingPrice', 0)

        # Create a DataFrame
        calls_df = pd.DataFrame(calls_data)

        # Ensure numeric types
        calls_df['strike'] = calls_df['strike'].astype(float)
        calls_df['bidPrice'] = calls_df['bidPrice'].astype(float)
        calls_df['askPrice'] = calls_df['askPrice'].astype(float)

        # Sort by strike price
        calls_df = calls_df.sort_values(by='strike')

        # Find ATM strike (closest to underlying price)
        atm_strike = calls_df.loc[(calls_df['strike'] - underlying_price).abs().idxmin(), 'strike']
        print(f"Underlying Price: {underlying_price}")
        print(f"ATM Strike: {atm_strike}")

        # Create initial base strike using global parameters
        base_strike = atm_strike + ATM_ADDER
        print(f"Initial Base Strike: {base_strike}")

        # Create all base strikes using global parameters
        base_strikes = [base_strike + i * INCREMENTAL_ADDER for i in range(NUM_BASES)]
        print(f"All Base Strikes: {base_strikes}")

        # Process each base strike
        current_result = {
            "symbol_name": symbol_name,
            "underlying_price": underlying_price,
            "atm_strike": atm_strike,
            "lot_size": lot_size,
            "trades": []
        }

        print(f"Processing {len(base_strikes)} base strikes with {len(OFFSET_VALUES)} offset values and {len(RATIO_SETS)} ratio sets...")

        for base in base_strikes:
            # Find closest available strike to the calculated base strike
            base_idx = (calls_df['strike'] - base).abs().idxmin()
            base_row = calls_df.loc[base_idx]
            base_strike_actual = base_row['strike']

            # For buying call options, use the ask price
            base_price = base_row['askPrice']
            base_instrument_key = base_row.get('instrument_key', 'N/A')

            print(f"Base strike: {base} -> Actual: {base_strike_actual}")

            # Process each offset value to create higher strikes
            for offset in OFFSET_VALUES:
                higher_strike = base + offset

                # Find closest available strike to the calculated higher strike
                higher_idx = (calls_df['strike'] - higher_strike).abs().idxmin()
                higher_row = calls_df.loc[higher_idx]
                higher_strike_actual = higher_row['strike']

                # Skip if strikes are the same (can happen with sparse option chains)
                if base_strike_actual == higher_strike_actual:
                    print(f"  Skipping offset {offset} - same as base strike")
                    continue

                # For selling call options, use the bid price
                higher_price = higher_row['bidPrice']
                higher_instrument_key = higher_row.get('instrument_key', 'N/A')

                print(f"  Higher strike (+{offset}): {higher_strike} -> Actual: {higher_strike_actual}")

                # Process each ratio set
                for buy_qty, sell_qty in RATIO_SETS:
                    # Calculate net premium
                    net_premium = (buy_qty * base_price) - (sell_qty * higher_price)

                    # Calculate PNL
                    pnl = net_premium * (-1) * lot_size

                    # Calculate percentage away from underlying
                    pct_away = ((base_strike_actual - underlying_price) / underlying_price) * 100

                    print(f"    Ratio {buy_qty}x{sell_qty}: Net Premium = {net_premium:.2f}, PNL = {pnl:.2f}")

                    # Create trade entry (include all trades, even if premium is not favorable)
                    trade = {
                        "base_strike": base_strike_actual,
                        "higher_strike": higher_strike_actual,
                        "buy_qty": buy_qty,
                        "sell_qty": sell_qty,
                        "buy_price": base_price,
                        "sell_price": higher_price,
                        "net_premium": round(net_premium, 2),
                        "pnl": round(pnl, 2),
                        "percentage_away": round(pct_away, 2),
                        "spread_width": higher_strike_actual - base_strike_actual,
                        "buy_instrument_key": base_instrument_key,
                        "sell_instrument_key": higher_instrument_key,
                        "strategy": f"{buy_qty}x{base_strike_actual}C / {sell_qty}x{higher_strike_actual}C",
                        "offset": offset,
                        "ratio": f"{buy_qty}x{sell_qty}"
                    }

                    # Add trade to results
                    current_result["trades"].append(trade)

        # Print summary of trades found
        print(f"Total trades found: {len(current_result['trades'])}")

        # Add current result to global results
        global_results.append(current_result)

        return current_result

    except Exception as e:
        print(f"Error analyzing option chain: {str(e)}")
        import traceback
        traceback.print_exc()
        return None

def sort_table(tree, col, descending):
    """Sort the table by the selected column"""
    data = [(tree.set(child, col), child) for child in tree.get_children("")]

    # Try converting to float for sorting if possible
    try:
        data.sort(key=lambda x: float(x[0]), reverse=descending)
    except ValueError:
        data.sort(reverse=descending)

    # Rearrange items in the tree
    for index, (_, child) in enumerate(data):
        tree.move(child, "", index)

    # Toggle sorting order for next click
    tree.heading(col, command=lambda: sort_table(tree, col, not descending))

def create_consolidated_display(parent_frame):
    """
    Create a consolidated display with all ratio sets in a single table
    """
    global global_results

    # Create a frame for the consolidated table
    table_frame = ttk.Frame(parent_frame)
    table_frame.pack(fill="both", expand=True, padx=10, pady=5)

    # Add ATM info header
    header_frame = ttk.Frame(table_frame)
    header_frame.pack(fill="x", pady=5)

    if global_results:
        result = global_results[0]
        atm_strike = result["atm_strike"]
        underlying_price = result.get("underlying_price", "N/A")

        ttk.Label(header_frame, text=f"Underlying: {underlying_price}", font=("Arial", 10, "bold")).grid(row=0, column=0, padx=5)
        ttk.Label(header_frame, text=f"ATM: {atm_strike}", font=("Arial", 10, "bold")).grid(row=0, column=1, padx=5)
        ttk.Label(header_frame, text=f"ATM adder: {ATM_ADDER}", font=("Arial", 10)).grid(row=0, column=2, padx=5)
        ttk.Label(header_frame, text=f"Adder Values: {', '.join(map(str, OFFSET_VALUES))}",
                  font=("Arial", 10)).grid(row=0, column=3, padx=5)

    # Create a canvas with scrollbars for the matrix
    canvas = tk.Canvas(table_frame)
    y_scrollbar = ttk.Scrollbar(table_frame, orient="vertical", command=canvas.yview)
    x_scrollbar = ttk.Scrollbar(table_frame, orient="horizontal", command=canvas.xview)

    canvas.configure(yscrollcommand=y_scrollbar.set, xscrollcommand=x_scrollbar.set)
    y_scrollbar.pack(side="right", fill="y")
    x_scrollbar.pack(side="bottom", fill="x")
    canvas.pack(side="left", fill="both", expand=True)

    matrix_frame = ttk.Frame(canvas)
    canvas.create_window((0, 0), window=matrix_frame, anchor="nw")

    # Determine how many ratio sets we have
    num_ratio_sets = len(RATIO_SETS)

    # Calculate total columns for each offset
    # Each ratio set will have Buy Price, Sell Price, Net Premium columns
    cols_per_ratio = 3
    cols_per_offset = num_ratio_sets * cols_per_ratio

    # Create main header row
    ttk.Label(matrix_frame, text="Base Strike", width=10, font=("Arial", 10, "bold"),
              borderwidth=1, relief="solid").grid(row=0, column=0, rowspan=3, padx=1, pady=1, sticky="nsew")

    # Create headers for each offset value with colspan
    col_offset = 1
    for offset in OFFSET_VALUES:
        # Column span includes all the ratio sets for this offset
        ttk.Label(matrix_frame, text=f"Offset: {offset}", width=cols_per_offset * 8, font=("Arial", 10, "bold"),
                  borderwidth=1, relief="solid").grid(row=0, column=col_offset, columnspan=cols_per_offset,
                                                      padx=1, pady=1, sticky="nsew")

        # Add ratio set headers for this offset
        for i, ratio_set in enumerate(RATIO_SETS):
            ratio_text = f"{ratio_set[0]}x{ratio_set[1]}"
            ratio_col = col_offset + (i * cols_per_ratio)
            ttk.Label(matrix_frame, text=ratio_text, width=24, font=("Arial", 10, "bold"),
                      borderwidth=1, relief="solid").grid(row=1, column=ratio_col,
                                                          columnspan=cols_per_ratio, padx=1, pady=1, sticky="nsew")

            # Add column headers for each ratio set
            ttk.Label(matrix_frame, text="Buy Price", width=8,
                      borderwidth=1, relief="solid").grid(row=2, column=ratio_col, padx=1, pady=1, sticky="nsew")
            ttk.Label(matrix_frame, text="Sell Price", width=8,
                      borderwidth=1, relief="solid").grid(row=2, column=ratio_col+1, padx=1, pady=1, sticky="nsew")
            ttk.Label(matrix_frame, text="Net Premium", width=8,
                      borderwidth=1, relief="solid").grid(row=2, column=ratio_col+2, padx=1, pady=1, sticky="nsew")

        col_offset += cols_per_offset

    # Create a data structure to organize all the data
    # Format: {base_strike: {offset: {ratio: {buy_price, sell_price, net_premium}}}}
    organized_data = {}

    if global_results:
        for result in global_results:
            for trade in result["trades"]:
                base_strike = trade["base_strike"]
                higher_strike = trade["higher_strike"]
                offset = higher_strike - base_strike
                buy_qty = trade["buy_qty"]
                sell_qty = trade["sell_qty"]
                ratio = f"{buy_qty}x{sell_qty}"

                # Skip if the offset is not in our predefined list
                if offset not in OFFSET_VALUES:
                    continue

                if base_strike not in organized_data:
                    organized_data[base_strike] = {}

                if offset not in organized_data[base_strike]:
                    organized_data[base_strike][offset] = {}

                organized_data[base_strike][offset][ratio] = {
                    "buy_price": trade["buy_price"],
                    "sell_price": trade["sell_price"],
                    "net_premium": trade["net_premium"]
                }

    # Sort base strikes
    base_strikes = sorted(organized_data.keys()) if organized_data else []

    # Fill in the table with data
    for row_idx, base_strike in enumerate(base_strikes):
        # Add the base strike column
        ttk.Label(matrix_frame, text=str(base_strike), width=10, font=("Arial", 10),
                  borderwidth=1, relief="solid").grid(row=row_idx+3, column=0, padx=1, pady=1, sticky="nsew")

        # Add data for each offset and ratio set
        col_offset = 1
        for offset in OFFSET_VALUES:
            offset_data = organized_data.get(base_strike, {}).get(offset, {})

            for i, ratio_set in enumerate(RATIO_SETS):
                ratio = f"{ratio_set[0]}x{ratio_set[1]}"
                ratio_col = col_offset + (i * cols_per_ratio)

                # Get data for this ratio or use placeholder
                if ratio in offset_data:
                    buy_price = offset_data[ratio]["buy_price"]
                    sell_price = offset_data[ratio]["sell_price"]
                    net_premium = offset_data[ratio]["net_premium"]

                    # Add color coding based on net premium value
                    if net_premium < 0:
                        premium_bg = "#ffcccc"  # Light red for negative
                    else:
                        premium_bg = "#ccffcc"  # Light green for positive
                else:
                    buy_price = "-"
                    sell_price = "-"
                    net_premium = "-"
                    premium_bg = "white"

                ttk.Label(matrix_frame, text=str(buy_price), width=8,
                          borderwidth=1, relief="solid").grid(row=row_idx+3, column=ratio_col,
                                                              padx=1, pady=1, sticky="nsew")
                ttk.Label(matrix_frame, text=str(sell_price), width=8,
                          borderwidth=1, relief="solid").grid(row=row_idx+3, column=ratio_col+1,
                                                              padx=1, pady=1, sticky="nsew")

                # For net premium, use a regular tk.Label to enable background color
                lbl = tk.Label(matrix_frame, text=str(net_premium), width=8, bg=premium_bg,
                               borderwidth=1, relief="solid")
                lbl.grid(row=row_idx+3, column=ratio_col+2, padx=1, pady=1, sticky="nsew")

            col_offset += cols_per_offset

    # Update scrollregion after adding all items
    matrix_frame.update_idletasks()
    canvas.config(scrollregion=canvas.bbox("all"))

    return table_frame

def create_top_premium_table(parent_frame):
    """Create a table showing the negative net premium results for each ratio set in ascending order"""
    global global_results

    # Create frame for the table
    table_frame = ttk.LabelFrame(parent_frame, text="Negative Net Premium Results (All Ratio Sets)")
    table_frame.pack(fill="both", expand=True, padx=10, pady=10)

    # Create treeview
    columns = ("ratio", "strategy", "base_strike", "higher_strike", "buy_price",
               "sell_price", "net_premium", "pnl", "percentage_away", "spread_width")

    tree = ttk.Treeview(table_frame, columns=columns, show="headings", height=10)

    # Define column headings
    tree.heading("ratio", text="Ratio", command=lambda: sort_table(tree, "ratio", False))
    tree.heading("strategy", text="Strategy", command=lambda: sort_table(tree, "strategy", False))
    tree.heading("base_strike", text="Base Strike", command=lambda: sort_table(tree, "base_strike", False))
    tree.heading("higher_strike", text="Higher Strike", command=lambda: sort_table(tree, "higher_strike", False))
    tree.heading("buy_price", text="Buy Price", command=lambda: sort_table(tree, "buy_price", False))
    tree.heading("sell_price", text="Sell Price", command=lambda: sort_table(tree, "sell_price", False))
    tree.heading("net_premium", text="Net Premium", command=lambda: sort_table(tree, "net_premium", False))
    tree.heading("pnl", text="PNL", command=lambda: sort_table(tree, "pnl", False))
    tree.heading("percentage_away", text="% Away", command=lambda: sort_table(tree, "percentage_away", False))
    tree.heading("spread_width", text="Spread Width", command=lambda: sort_table(tree, "spread_width", False))

    # Define column widths
    tree.column("ratio", width=60, anchor="center")
    tree.column("strategy", width=150, anchor="center")
    tree.column("base_strike", width=80, anchor="center")
    tree.column("higher_strike", width=80, anchor="center")
    tree.column("buy_price", width=70, anchor="center")
    tree.column("sell_price", width=70, anchor="center")
    tree.column("net_premium", width=100, anchor="center")
    tree.column("pnl", width=100, anchor="center")
    tree.column("percentage_away", width=70, anchor="center")
    tree.column("spread_width", width=80, anchor="center")

    # Add scrollbar
    scrollbar = ttk.Scrollbar(table_frame, orient="vertical", command=tree.yview)
    scrollbar_h = ttk.Scrollbar(table_frame, orient="horizontal", command=tree.xview)
    tree.configure(yscrollcommand=scrollbar.set, xscrollcommand=scrollbar_h.set)

    scrollbar.pack(side="right", fill="y")
    scrollbar_h.pack(side="bottom", fill="x")
    tree.pack(side="left", fill="both", expand=True)

    # Populate table if data exists
    if global_results:
        all_trades = []
        for result in global_results:
            all_trades.extend(result["trades"])

        # Filter only trades with negative net premium
        negative_premium_trades = [trade for trade in all_trades if trade["net_premium"] < 0]

        # Group by ratio
        negative_trades_per_ratio = {}
        for trade in negative_premium_trades:
            ratio = f"{trade['buy_qty']}x{trade['sell_qty']}"
            if ratio not in negative_trades_per_ratio:
                negative_trades_per_ratio[ratio] = []
            negative_trades_per_ratio[ratio].append(trade)

        # Sort each ratio's trades by net premium (ascending) and get top 3 most negative
        all_negative_trades = []
        for ratio, trades in negative_trades_per_ratio.items():
            sorted_trades = sorted(trades, key=lambda x: x["net_premium"])  # Ascending order (most negative first)
            all_negative_trades.extend(sorted_trades[:3])  # Get top 3 most negative trades per ratio

        # Sort all trades by net premium (ascending)
        sorted_negative_trades = sorted(all_negative_trades, key=lambda x: x["net_premium"])

        # Add to treeview
        for i, trade in enumerate(sorted_negative_trades):
            tree.insert("", "end", values=(
                f"{trade['buy_qty']}x{trade['sell_qty']}",
                trade["strategy"],
                trade["base_strike"],
                trade["higher_strike"],
                trade["buy_price"],
                trade["sell_price"],
                trade["net_premium"],
                trade["pnl"],
                trade["percentage_away"],
                trade["spread_width"]
            ))

        # Sort initially by net premium (ascending - most negative first)
        sort_table(tree, "net_premium", False)

    return table_frame

def create_application_ui(root):
    """Create the main application UI"""
    global status_text

    # Create top frame for fetch button and status
    top_frame = ttk.Frame(root, padding=10)
    top_frame.pack(fill="x", pady=5)

    # Add note about API tokens
    ttk.Label(top_frame, text="Note: Make sure API tokens are valid and expiry dates use YYYY-MM-DD format",
              font=("Arial", 10, "italic")).pack(anchor="w", pady=5)

    # Add status label
    status_label = ttk.Label(top_frame, text="Ready")
    status_label.pack(fill="x", pady=5)
    status_text = status_label

    # Add fetch button
    fetch_button = ttk.Button(top_frame, text="Fetch & Analyze Data",
                              command=lambda: fetch_and_analyze(root, status_label))
    fetch_button.pack(pady=5)

    # Create main frame
    main_frame = ttk.Frame(root)
    main_frame.pack(fill="both", expand=True, padx=10, pady=5)

    # Use the new consolidated display instead of the tabbed display
    create_consolidated_display(main_frame)

    # Create top premium table below matrix display
    create_top_premium_table(root)

def fetch_and_analyze(root, status_label):
    """Fetch and analyze data for all instruments"""
    global global_results
    global_results = []

    status_label.config(text="Fetching and analyzing data... Please wait")
    root.update()

    try:
        # Process all instruments
        for instrument in INSTRUMENTS:
            instrument_key = instrument["instrument_key"]
            expiry_date = instrument["expiry_date"]
            symbol_name = instrument["symbol_name"]
            lot_size = instrument["lot_size"]

            status_label.config(text=f"Fetching data for {symbol_name}...")
            root.update()

            # Fetch option chain data
            option_chain = fetch_option_chain(instrument_key, expiry_date)

            if option_chain:
                status_label.config(text=f"Analyzing {symbol_name}...")
                root.update()

                # Analyze option chain data
                analyze_option_chain(option_chain, symbol_name, lot_size)
            else:
                status_label.config(text=f"Failed to fetch data for {symbol_name}")
                root.update()
                time.sleep(1)

        # Recreate the display with the new data
        for widget in root.winfo_children():
            widget.destroy()

        # Rebuild UI with new data
        create_application_ui(root)

    except Exception as e:
        status_label.config(text=f"Error: {str(e)}")
        print(f"Error fetching/analyzing data: {str(e)}")
        import traceback
        traceback.print_exc()

def main():
    """Main function"""
    root = tk.Tk()
    root.title("Call Option Strategy Analysis")
    root.geometry("1400x900")  # Increased size for wider table

    # Print global strategy parameters
    print("Strategy Parameters:")
    print(f"ATM Adder: {ATM_ADDER}")
    print(f"Incremental Adder: {INCREMENTAL_ADDER}")
    print(f"Number of Base Strikes: {NUM_BASES}")
    print(f"Offset Values: {OFFSET_VALUES}")
    print(f"Ratio Sets: {RATIO_SETS}")
    print(f"Number of Instruments: {len(INSTRUMENTS)}")
    print(f"Number of API Tokens: {len(API_TOKENS)}")

    # Create application UI
    create_application_ui(root)

    # Start the main loop
    root.mainloop()

if __name__ == "__main__":
    main()