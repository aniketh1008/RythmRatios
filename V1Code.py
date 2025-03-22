

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

# Global strategy parameters
ATM_ADDER = 150
INCREMENTAL_ADDER = 50
NUM_BASES = 4
OFFSET_VALUES = [250, 300, 350]
RATIO_SETS = [[1, 5], [1, 4], [2, 7], [1, 3], [2, 5], [3, 7], [1, 2], [4, 7], [2, 3]]

# API and symbol configuration
API_TOKENS = [
    "eyJ0eXAiOiJKV1QiLCJrZXlfaWQiOiJza192MS4wIiwiYWxnIjoiSFMyNTYifQ.eyJzdWIiOiJBSjQ5NzIiLCJqdGkiOiI2N2RkOWI1YjkyNzAyZjE3MjgwYjhmMmQiLCJpc011bHRpQ2xpZW50IjpmYWxzZSwiaWF0IjoxNzQyNTc2NDc1LCJpc3MiOiJ1ZGFwaS1nYXRld2F5LXNlcnZpY2UiLCJleHAiOjE3NDI1OTQ0MDB9.UYm-hguxfsqSho8e8NI6xYCvMxS4G-RH9uUaUpf4R10",
    "eyJ0eXAiOiJKV1QiLCJrZXlfaWQiOiJza192MS4wIiwiYWxnIjoiSFMyNTYifQ.eyJzdWIiOiIzSkNZTUciLCJqdGkiOiI2N2RkOWU0MDkyNzAyZjE3MjgwYjhmNGYiLCJpc011bHRpQ2xpZW50IjpmYWxzZSwiaWF0IjoxNzQyNTc3MjE2LCJpc3MiOiJ1ZGFwaS1nYXRld2F5LXNlcnZpY2UiLCJleHAiOjE3NDI1OTQ0MDB9.yQB4rO8bTcTY_p17LloxRK1jjvXcsexrOkg95XvT4X8"
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
        Expiry date (format: "2025-03-27")

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

        for base in base_strikes:
            # Find closest available strike to the calculated base strike
            base_row = calls_df.loc[(calls_df['strike'] - base).abs().idxmin()]
            base_strike_actual = base_row['strike']

            # For buying call options, use the ask price
            base_price = base_row['askPrice']
            base_instrument_key = base_row.get('instrument_key', 'N/A')

            # Process each offset value to create higher strikes
            for offset in OFFSET_VALUES:
                higher_strike = base + offset

                # Find closest available strike to the calculated higher strike
                higher_row = calls_df.loc[(calls_df['strike'] - higher_strike).abs().idxmin()]
                higher_strike_actual = higher_row['strike']

                # Skip if strikes are the same (can happen with sparse option chains)
                if base_strike_actual == higher_strike_actual:
                    continue

                # For selling call options, use the bid price
                higher_price = higher_row['bidPrice']
                higher_instrument_key = higher_row.get('instrument_key', 'N/A')

                # Process each ratio set
                for buy_qty, sell_qty in RATIO_SETS:
                    # Calculate net premium
                    net_premium = (buy_qty * base_price) - (sell_qty * higher_price)

                    # Skip if premium is not favorable (we want to receive premium, not pay it)
                    if net_premium >= 0:
                        continue

                    # Calculate PNL
                    pnl = net_premium * (-1) * lot_size

                    # Calculate percentage away from underlying
                    pct_away = ((base_strike_actual - underlying_price) / underlying_price) * 100

                    # Create trade entry
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
                        "strategy": f"{buy_qty}x{base_strike_actual}C / {sell_qty}x{higher_strike_actual}C"
                    }

                    # Add trade to results
                    current_result["trades"].append(trade)

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

def display_results_table(root):
    """
    Display results in a Tkinter table with the desired formatting
    """
    global global_results

    notebook = ttk.Notebook(root)
    notebook.pack(fill='both', expand=True, padx=10, pady=10)

    # Create a tab for each ratio set
    for buy_qty, sell_qty in RATIO_SETS:
        ratio_key = f"{buy_qty}x{sell_qty}"

        # Create a frame for this ratio
        tab_frame = ttk.Frame(notebook)
        notebook.add(tab_frame, text=f"Ratio {ratio_key}")

        # Create a frame for info at the top
        info_frame = ttk.Frame(tab_frame)
        info_frame.pack(fill='x', padx=5, pady=5)

        for result in global_results:
            symbol_name = result['symbol_name']
            atm_strike = result['atm_strike']

            # Only display info for the first result
            ttk.Label(info_frame, text=f"Symbol: {symbol_name}").grid(row=0, column=0, sticky='w', padx=5)
            ttk.Label(info_frame, text=f"ATM: {atm_strike}").grid(row=0, column=1, sticky='w', padx=5)
            ttk.Label(info_frame, text=f"ATM Strike Adder: {ATM_ADDER}").grid(row=0, column=2, sticky='w', padx=5)
            ttk.Label(info_frame, text=f"Adder Values: {', '.join(map(str, OFFSET_VALUES))}").grid(row=0, column=3, sticky='w', padx=5)

            # We just need one of these, so break after the first
            break

        # Create a frame for the scrollable table area
        table_frame = ttk.Frame(tab_frame)
        table_frame.pack(fill='both', expand=True, padx=5, pady=5)

        # Add vertical scrollbar
        vsb = ttk.Scrollbar(table_frame, orient="vertical")
        vsb.pack(side='right', fill='y')

        # Add horizontal scrollbar
        hsb = ttk.Scrollbar(table_frame, orient="horizontal")
        hsb.pack(side='bottom', fill='x')

        # Define columns for the table - will set dynamically based on offset values
        # First column is the base strike, followed by pairs of [higher_strike, net_premium] for each offset
        columns = ["Base Strike"] + [f"Higher Strike ({offset})" for offset in OFFSET_VALUES] + [f"Net Premium ({offset})" for offset in OFFSET_VALUES]

        # Create Treeview
        tree = ttk.Treeview(table_frame, columns=columns, show="headings",
                            yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        vsb.config(command=tree.yview)
        hsb.config(command=tree.xview)

        # Configure column headings
        for i, col in enumerate(columns):
            tree.heading(col, text=col)
            tree.column(col, width=120, anchor="center")

        tree.pack(side='left', fill='both', expand=True)

        # Now populate the table with data
        for result in global_results:
            # Group trades by base_strike for this ratio
            base_strike_groups = {}

            for trade in result['trades']:
                if trade['buy_qty'] == buy_qty and trade['sell_qty'] == sell_qty:
                    base_strike = trade['base_strike']
                    higher_strike = trade['higher_strike']
                    net_premium = trade['net_premium']

                    # Calculate which offset was used
                    offset = higher_strike - base_strike

                    # Initialize if not already present
                    if base_strike not in base_strike_groups:
                        base_strike_groups[base_strike] = {}

                    # Store higher strike and net premium for this offset
                    base_strike_groups[base_strike][offset] = {
                        'higher_strike': higher_strike,
                        'net_premium': net_premium
                    }

            # Now add rows to the tree
            for base_strike, offsets in sorted(base_strike_groups.items()):
                # Start with the base strike
                row_values = [base_strike]

                # Add higher strike values
                for offset in OFFSET_VALUES:
                    if offset in offsets:
                        row_values.append(offsets[offset]['higher_strike'])
                    else:
                        row_values.append("-")

                # Add net premium values
                for offset in OFFSET_VALUES:
                    if offset in offsets:
                        row_values.append(offsets[offset]['net_premium'])
                    else:
                        row_values.append("-")

                # Insert row into tree
                tree.insert("", "end", values=row_values)

        # Add the export button at the bottom of the tab
        button_frame = ttk.Frame(tab_frame)
        button_frame.pack(fill='x', padx=5, pady=5)

        def export_to_excel(ratio=ratio_key):
            from datetime import datetime
            filename = f"call_options_{ratio}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"

            # Create DataFrame from tree data
            data = []

            for item in tree.get_children():
                values = tree.item(item)["values"]
                data.append(values)

            df = pd.DataFrame(data, columns=columns)
            df.to_excel(filename, index=False)
            messagebox.showinfo("Export Complete", f"Data exported to {filename}")

        export_btn = ttk.Button(button_frame, text=f"Export {ratio_key} to Excel",
                                command=lambda r=ratio_key: export_to_excel(r))
        export_btn.pack(side='right', padx=5)

    # Add a tab for analyzing
    analyze_tab = ttk.Frame(notebook)
    notebook.add(analyze_tab, text="Fetch & Analyze")

    # Add controls for fetching data
    control_frame = ttk.Frame(analyze_tab, padding=10)
    control_frame.pack(fill='x')

    status_frame = ttk.Frame(analyze_tab, padding=10)
    status_frame.pack(fill='x')

    status_label = ttk.Label(status_frame, text="")
    status_label.pack(fill='x')

    # Add fetch and analyze all button
    def fetch_and_analyze_all():
        # Clear previous results
        global global_results
        global_results = []
        status_label.config(text="Fetching and analyzing all instruments... Please wait")
        root.update()  # Force update to show status message

        try:
            # Process each instrument
            for instrument in INSTRUMENTS:
                instrument_key = instrument["instrument_key"]
                expiry_date = instrument["expiry_date"]
                symbol_name = instrument["symbol_name"]
                lot_size = instrument["lot_size"]

                status_label.config(text=f"Fetching data for {symbol_name}... Please wait")
                root.update()

                # Fetch option chain
                option_chain = fetch_option_chain(instrument_key, expiry_date)

                if option_chain:
                    status_label.config(text=f"Analyzing {symbol_name}... Please wait")
                    root.update()

                    # Analyze the data
                    analyze_option_chain(option_chain, symbol_name, lot_size)

            # Update all tabs with new data
            # This requires recreating the notebook, so we'll temporarily store the current tab
            current_tab = notebook.index(notebook.select())
            notebook.forget(0)  # Remove all tabs

            # Rebuild the notebook with the new data
            display_results_table(root)

            # Try to restore the previously selected tab
            if current_tab < notebook.index("end"):
                notebook.select(current_tab)

            status_label.config(text=f"Analysis completed for {len(INSTRUMENTS)} instruments")

        except Exception as e:
            status_label.config(text=f"Error during analysis: {str(e)}")
            import traceback
            traceback.print_exc()

    # Add the fetch all button
    fetch_btn = ttk.Button(control_frame, text="Fetch & Analyze All Instruments",
                           command=fetch_and_analyze_all, style="Accent.TButton")
    fetch_btn.pack(side='left', padx=10)

    # Create a style for highlighted button
    style = ttk.Style()
    style.configure("Accent.TButton", foreground="black", background="#4CAF50", font=('Helvetica', 10, 'bold'))

def main():
    """Main function"""
    root = tk.Tk()
    root.title("Call Option Strategy Analysis")
    root.geometry("1400x800")

    # Print global strategy parameters
    print("Strategy Parameters:")
    print(f"ATM Adder: {ATM_ADDER}")
    print(f"Incremental Adder: {INCREMENTAL_ADDER}")
    print(f"Number of Base Strikes: {NUM_BASES}")
    print(f"Offset Values: {OFFSET_VALUES}")
    print(f"Ratio Sets: {RATIO_SETS}")
    print(f"Number of Instruments: {len(INSTRUMENTS)}")
    print(f"Number of API Tokens: {len(API_TOKENS)}")

    # Display results table
    display_results_table(root)

    # Start the main loop
    root.mainloop()

if __name__ == "__main__":
    main()