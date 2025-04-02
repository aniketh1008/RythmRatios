
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

# Global strategy parameters - These will be overridden by UI inputs
ATM_ADDER = 150
INCREMENTAL_ADDER = 50
NUM_BASES = 8
OFFSET_VALUES = [250, 300, 350]
RATIO_SETS = [[1, 5], [1, 4], [2, 7], [1, 3], [2, 5], [3, 7], [1, 2], [4, 7], [2, 3]]

# API and symbol configuration
API_TOKENS = [
    "eyJ0eXAiOiJKV1QiLCJrZXlfaWQiOiJza192MS4wIiwiYWxnIjoiSFMyNTYifQ.eyJzdWIiOiJBSjQ5NzIiLCJqdGkiOiI2N2ViNTY2Zjc5NzljNDA1OTkzODUyZGEiLCJpc011bHRpQ2xpZW50IjpmYWxzZSwiaWF0IjoxNzQzNDc2MzM1LCJpc3MiOiJ1ZGFwaS1nYXRld2F5LXNlcnZpY2UiLCJleHAiOjE3NDM1NDQ4MDB9.uugONCxiSNWF7USiKhXD8G2kOx9lrC4gwmYQ95crlXA",
    "eyJ0eXAiOiJKV1QiLCJrZXlfaWQiOiJza192MS4wIiwiYWxnIjoiSFMyNTYifQ.eyJzdWIiOiIzSkNZTUciLCJqdGkiOiI2N2ViNTY5YjdkNDNmNDNhZTY1YWY2ZmIiLCJpc011bHRpQ2xpZW50IjpmYWxzZSwiaWF0IjoxNzQzNDc2Mzc5LCJpc3MiOiJ1ZGFwaS1nYXRld2F5LXNlcnZpY2UiLCJleHAiOjE3NDM1NDQ4MDB9.Fuka-5SDTM-GWWh4SnjVFvXNjoZzW4nVqvh1lNAh5mQ"
]

# List of instruments to analyze
INSTRUMENTS = [
    {
        "instrument_key": "NSE_INDEX|Nifty 50",
        "symbol_name": "NIFTY",
        "lot_size": 75,
        "expiry_date": "2025-04-03"
    }
    # Add more instruments as needed
]

# Global variables
global_results = {
    "calls": [],  # Stores analyzed call option data
    "puts": []    # Stores analyzed put option data
}
status_text = None  # Reference to status label
fetch_button = None  # Reference to fetch button
auto_poll_job = None  # Reference to polling job

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

# Add these new functions after your global variables

# Functions to check if a call/put spread fits the curve
def check_call_curve_fit(option_chain_data, atm_strike, base_strike, higher_strike):
    """Check if a call ratio spread fits the curve by comparing price ratios"""
    try:
        # Extract call options data
        calls_data = []
        for item in option_chain_data["data"]:
            if "call_options" in item and "market_data" in item["call_options"]:
                market_data = item["call_options"]["market_data"]
                calls_data.append({
                    "strike": item["strike_price"],
                    "bidPrice": market_data["bid_price"],
                    "askPrice": market_data["ask_price"]
                })

        # Create DataFrame
        calls_df = pd.DataFrame(calls_data)
        calls_df['strike'] = calls_df['strike'].astype(float)
        calls_df['bidPrice'] = calls_df['bidPrice'].astype(float)
        calls_df['askPrice'] = calls_df['askPrice'].astype(float)

        # Find data for strikes
        atm_row = calls_df.loc[(calls_df['strike'] - atm_strike).abs().idxmin()]
        base_row = calls_df.loc[(calls_df['strike'] - base_strike).abs().idxmin()]
        higher_row = calls_df.loc[(calls_df['strike'] - higher_strike).abs().idxmin()]

        # Calculate ratios
        ratio1 = atm_row['bidPrice'] / base_row['askPrice']
        ratio2 = base_row['bidPrice'] / higher_row['askPrice']

        return ratio1 > ratio2
    except Exception as e:
        print(f"Error in check_call_curve_fit: {str(e)}")
        return False

def check_put_curve_fit(option_chain_data, atm_strike, base_strike, lower_strike):
    """Check if a put ratio spread fits the curve by comparing price ratios"""
    try:
        # Extract put options data
        puts_data = []
        for item in option_chain_data["data"]:
            if "put_options" in item and "market_data" in item["put_options"]:
                market_data = item["put_options"]["market_data"]
                puts_data.append({
                    "strike": item["strike_price"],
                    "bidPrice": market_data["bid_price"],
                    "askPrice": market_data["ask_price"]
                })

        # Create DataFrame
        puts_df = pd.DataFrame(puts_data)
        puts_df['strike'] = puts_df['strike'].astype(float)
        puts_df['bidPrice'] = puts_df['bidPrice'].astype(float)
        puts_df['askPrice'] = puts_df['askPrice'].astype(float)

        # Calculate strike difference
        strike_diff = base_strike - lower_strike

        # Find data for strikes
        atm_row = puts_df.loc[(puts_df['strike'] - atm_strike).abs().idxmin()]
        base_row = puts_df.loc[(puts_df['strike'] - base_strike).abs().idxmin()]
        lower_row = puts_df.loc[(puts_df['strike'] - lower_strike).abs().idxmin()]

        # Find corresponding higher strike (ATM - strike_diff)
        temp_higher_strike = atm_strike - strike_diff
        temp_higher_row = puts_df.loc[(puts_df['strike'] - temp_higher_strike).abs().idxmin()]

        # Calculate ratios
        ratio1 = atm_row['bidPrice'] / temp_higher_row['askPrice']
        ratio2 = base_row['bidPrice'] / lower_row['askPrice']

        return ratio1 > ratio2
    except Exception as e:
        print(f"Error in check_put_curve_fit: {str(e)}")
        return False



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

def analyze_call_option_chain(option_chain_data, symbol_name="NIFTY", lot_size=50, strategy_params=None, check_curve_fit=False):
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
    strategy_params : dict
        Strategy parameters from UI (atm_adder, incremental_adder, num_bases, offset_values)

    Returns:
    --------
    dict
        Results of the analysis
    """
    try:
        global global_results

        # Use strategy parameters from UI if provided, otherwise use defaults
        if strategy_params:
            atm_adder = strategy_params['atm_adder']
            incremental_adder = strategy_params['incremental_adder']
            num_bases = strategy_params['num_bases']
            offset_values = strategy_params['offset_values']
        else:
            atm_adder = ATM_ADDER
            incremental_adder = INCREMENTAL_ADDER
            num_bases = NUM_BASES
            offset_values = OFFSET_VALUES

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

        # Create initial base strike using provided parameters
        base_strike = atm_strike + atm_adder
        print(f"Initial Base Strike: {base_strike}")

        # Create all base strikes using provided parameters
        base_strikes = [base_strike + i * incremental_adder for i in range(num_bases)]
        print(f"All Base Strikes: {base_strikes}")

        # Process each base strike
        current_result = {
            "symbol_name": symbol_name,
            "underlying_price": underlying_price,
            "atm_strike": atm_strike,
            "lot_size": lot_size,
            "option_chain_data": option_chain_data,
            "strategy_params": {
                "atm_adder": atm_adder,
                "incremental_adder": incremental_adder,
                "num_bases": num_bases,
                "offset_values": offset_values
            },
            "trades": []
        }

        print(f"Processing {len(base_strikes)} base strikes with {len(offset_values)} offset values and {len(RATIO_SETS)} ratio sets...")

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
            for offset in offset_values:
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

                # Check if the spread fits the curve if required
                if check_curve_fit:
                    curve_fit = check_call_curve_fit(option_chain_data, atm_strike, base_strike_actual, higher_strike_actual)
                    if curve_fit:
                        print("Curve is Fit")
                    if not curve_fit:
                        print(f"  Skipping spread - does not fit the curve")
                        continue

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
        print(f"Total call trades found: {len(current_result['trades'])}")

        # Add current result to global results
        global_results["calls"].append(current_result)

        return current_result

    except Exception as e:
        print(f"Error analyzing call option chain: {str(e)}")
        import traceback
        traceback.print_exc()
        return None

def analyze_put_option_chain(option_chain_data, symbol_name="NIFTY", lot_size=50, strategy_params=None, check_curve_fit=False):
    """
    Analyze put options from option chain data for put ratio spreads

    Parameters:
    -----------
    option_chain_data : dict
        Option chain data containing puts
    symbol_name : str
        Name of the symbol
    lot_size : int
        Lot size for the symbol
    strategy_params : dict
        Strategy parameters from UI (atm_adder, incremental_adder, num_bases, offset_values)

    Returns:
    --------
    dict
        Results of the analysis
    """
    try:
        global global_results

        # Use strategy parameters from UI if provided, otherwise use defaults
        if strategy_params:
            atm_adder = strategy_params['atm_adder']
            incremental_adder = strategy_params['incremental_adder']
            num_bases = strategy_params['num_bases']
            offset_values = strategy_params['offset_values']
        else:
            atm_adder = ATM_ADDER
            incremental_adder = INCREMENTAL_ADDER
            num_bases = NUM_BASES
            offset_values = OFFSET_VALUES

        # Check if data is in Upstox API format
        if isinstance(option_chain_data, dict) and "data" in option_chain_data:
            # Extract underlying price
            underlying_price = option_chain_data["data"][0]["underlying_spot_price"]

            # Extract put options data
            puts_data = []
            for item in option_chain_data["data"]:
                if "put_options" in item and "market_data" in item["put_options"]:
                    market_data = item["put_options"]["market_data"]
                    # Skip if bid or ask is not available
                    if market_data["bid_price"] <= 0 or market_data["ask_price"] <= 0:
                        continue

                    puts_data.append({
                        "strike": item["strike_price"],
                        "bidPrice": market_data["bid_price"],
                        "askPrice": market_data["ask_price"],
                        "instrument_key": item["put_options"]["instrument_key"]
                    })
        else:
            # Use sample data format
            puts_data = option_chain_data.get('puts', [])
            underlying_price = option_chain_data.get('underlyingPrice', 0)

        # Create a DataFrame
        puts_df = pd.DataFrame(puts_data)

        # Ensure numeric types
        puts_df['strike'] = puts_df['strike'].astype(float)
        puts_df['bidPrice'] = puts_df['bidPrice'].astype(float)
        puts_df['askPrice'] = puts_df['askPrice'].astype(float)

        # Sort by strike price
        puts_df = puts_df.sort_values(by='strike')

        # Find ATM strike (closest to underlying price)
        atm_strike = puts_df.loc[(puts_df['strike'] - underlying_price).abs().idxmin(), 'strike']
        print(f"Underlying Price: {underlying_price}")
        print(f"ATM Strike for Puts: {atm_strike}")

        # For puts, we want to go in the opposite direction from calls
        # Lower strikes have higher premiums (opposite of calls)
        base_strike = atm_strike - atm_adder
        print(f"Initial Base Strike for Puts: {base_strike}")

        # Create all base strikes using provided parameters, moving downward
        base_strikes = [base_strike - i * incremental_adder for i in range(num_bases)]
        print(f"All Base Strikes for Puts: {base_strikes}")

        # Process each base strike
        current_result = {
            "symbol_name": symbol_name,
            "underlying_price": underlying_price,
            "atm_strike": atm_strike,
            "lot_size": lot_size,
            "strategy_params": {
                "atm_adder": atm_adder,
                "incremental_adder": incremental_adder,
                "num_bases": num_bases,
                "offset_values": offset_values
            },
            "trades": []
        }

        print(f"Processing {len(base_strikes)} put base strikes with {len(offset_values)} offset values and {len(RATIO_SETS)} ratio sets...")

        for base in base_strikes:
            # Find closest available strike to the calculated base strike
            base_idx = (puts_df['strike'] - base).abs().idxmin()
            base_row = puts_df.loc[base_idx]
            base_strike_actual = base_row['strike']

            # For buying put options, use the ask price
            base_price = base_row['askPrice']
            base_instrument_key = base_row.get('instrument_key', 'N/A')

            print(f"Base put strike: {base} -> Actual: {base_strike_actual}")

            # Process each offset value to create lower strikes
            for offset in offset_values:
                lower_strike = base - offset

                # Find closest available strike to the calculated lower strike
                lower_idx = (puts_df['strike'] - lower_strike).abs().idxmin()

                # Skip if index is out of range (can happen with small option chains)
                if lower_idx not in puts_df.index:
                    print(f"  Skipping offset {offset} - strike not found")
                    continue

                lower_row = puts_df.loc[lower_idx]
                lower_strike_actual = lower_row['strike']

                # Skip if strikes are the same (can happen with sparse option chains)
                if base_strike_actual == lower_strike_actual:
                    print(f"  Skipping offset {offset} - same as base strike")
                    continue

                # For selling put options, use the bid price
                lower_price = lower_row['bidPrice']
                lower_instrument_key = lower_row.get('instrument_key', 'N/A')

                print(f"  Lower strike (-{offset}): {lower_strike} -> Actual: {lower_strike_actual}")

                # Check if the spread fits the curve if required
                if check_curve_fit:
                    curve_fit = check_put_curve_fit(option_chain_data, atm_strike, base_strike_actual, lower_strike_actual)
                    if curve_fit:
                        print("Curve is Fit")
                    if not curve_fit:
                        print(f"  Skipping spread - does not fit the curve")
                        continue

                # Process each ratio set
                for buy_qty, sell_qty in RATIO_SETS:
                    # Calculate net premium
                    net_premium = (buy_qty * base_price) - (sell_qty * lower_price)

                    # Calculate PNL
                    pnl = net_premium * (-1) * lot_size

                    # Calculate percentage away from underlying
                    pct_away = ((underlying_price - base_strike_actual) / underlying_price) * 100

                    print(f"    Ratio {buy_qty}x{sell_qty}: Net Premium = {net_premium:.2f}, PNL = {pnl:.2f}")

                    # Create trade entry (include all trades, even if premium is not favorable)
                    trade = {
                        "base_strike": base_strike_actual,
                        "lower_strike": lower_strike_actual,
                        "buy_qty": buy_qty,
                        "sell_qty": sell_qty,
                        "buy_price": base_price,
                        "sell_price": lower_price,
                        "net_premium": round(net_premium, 2),
                        "pnl": round(pnl, 2),
                        "percentage_away": round(pct_away, 2),
                        "spread_width": base_strike_actual - lower_strike_actual,
                        "buy_instrument_key": base_instrument_key,
                        "sell_instrument_key": lower_instrument_key,
                        "strategy": f"{buy_qty}x{base_strike_actual}P / {sell_qty}x{lower_strike_actual}P",
                        "offset": offset,
                        "ratio": f"{buy_qty}x{sell_qty}"
                    }

                    # Add trade to results
                    current_result["trades"].append(trade)

        # Print summary of trades found
        print(f"Total put trades found: {len(current_result['trades'])}")

        # Add current result to global results
        global_results["puts"].append(current_result)

        return current_result

    except Exception as e:
        print(f"Error analyzing put option chain: {str(e)}")
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

def create_consolidated_display(parent_frame, option_type="calls"):
    """
    Create separate tables for each offset value (from UI parameters)
    """
    global global_results

    # Create a frame to hold all three table frames
    main_display_frame = ttk.Frame(parent_frame)
    main_display_frame.pack(fill="both", expand=True, padx=5, pady=5)

    # Add ATM info header at the top
    header_frame = ttk.Frame(main_display_frame)
    header_frame.pack(fill="x", pady=5)

    results_list = global_results.get(option_type, [])
    if results_list:
        result = results_list[0]
        atm_strike = result["atm_strike"]
        underlying_price = result.get("underlying_price", "N/A")
        strategy_params = result.get("strategy_params", {})

        atm_adder = strategy_params.get("atm_adder", ATM_ADDER)
        incremental_adder = strategy_params.get("incremental_adder", INCREMENTAL_ADDER)

        option_type_label = "Calls" if option_type == "calls" else "Puts"
        ttk.Label(header_frame, text=f"{option_type_label} - Underlying: {underlying_price}", font=("Arial", 10, "bold")).grid(row=0, column=0, padx=5)
        ttk.Label(header_frame, text=f"ATM: {atm_strike}", font=("Arial", 10, "bold")).grid(row=0, column=1, padx=5)
        ttk.Label(header_frame, text=f"ATM adder: {atm_adder}", font=("Arial", 10)).grid(row=0, column=2, padx=5)
        ttk.Label(header_frame, text=f"Incremental adder: {incremental_adder}", font=("Arial", 10)).grid(row=0, column=3, padx=5)

    # Create a notebook to hold the tables
    notebook = ttk.Notebook(main_display_frame)
    notebook.pack(fill="both", expand=True, padx=5, pady=5)

    # Create organized data structure
    organized_data = organize_trade_data(option_type)

    # Get the offset values from the first result (if available)
    if results_list:
        strategy_params = results_list[0].get("strategy_params", {})
        offset_values = strategy_params.get("offset_values", OFFSET_VALUES)
    else:
        offset_values = OFFSET_VALUES

    # Create a separate table for each offset value
    for offset in offset_values:
        if option_type == "calls":
            create_call_offset_table(notebook, offset, organized_data)
        else:
            create_put_offset_table(notebook, offset, organized_data)

    return main_display_frame

def organize_trade_data(option_type="calls"):
    """Organize trade data for table display"""
    global global_results

    # Create a data structure to organize all the data
    # Format: {offset: {base_strike: {ratio: {buy_price, sell_price, net_premium}}}}
    organized_data = {}

    results_list = global_results.get(option_type, [])
    if results_list:
        for result in results_list:
            for trade in result["trades"]:
                base_strike = trade["base_strike"]

                if option_type == "calls":
                    other_strike = trade["higher_strike"]
                    offset = other_strike - base_strike
                else:  # puts
                    other_strike = trade["lower_strike"]
                    offset = base_strike - other_strike

                buy_qty = trade["buy_qty"]
                sell_qty = trade["sell_qty"]
                ratio = f"{buy_qty}x{sell_qty}"

                # Get the offset values from the result's strategy parameters
                strategy_params = result.get("strategy_params", {})
                offset_values = strategy_params.get("offset_values", OFFSET_VALUES)

                # Skip if the offset is not in our predefined list
                if offset not in offset_values:
                    continue

                if offset not in organized_data:
                    organized_data[offset] = {}

                if base_strike not in organized_data[offset]:
                    organized_data[offset][base_strike] = {}

                if ratio not in organized_data[offset][base_strike]:
                    organized_data[offset][base_strike][ratio] = {
                        "buy_price": trade["buy_price"],
                        "sell_price": trade["sell_price"],
                        "net_premium": trade["net_premium"]
                    }

    return organized_data

def create_call_offset_table(notebook, offset, organized_data):
    """Create a table for a specific call offset value"""
    # Create a frame for this offset
    offset_frame = ttk.Frame(notebook)
    notebook.add(offset_frame, text=f"Offset: +{offset}")

    # Create a canvas with scrollbars for the matrix
    canvas = tk.Canvas(offset_frame)
    y_scrollbar = ttk.Scrollbar(offset_frame, orient="vertical", command=canvas.yview)
    x_scrollbar = ttk.Scrollbar(offset_frame, orient="horizontal", command=canvas.xview)

    canvas.configure(yscrollcommand=y_scrollbar.set, xscrollcommand=x_scrollbar.set)
    y_scrollbar.pack(side="right", fill="y")
    x_scrollbar.pack(side="bottom", fill="x")
    canvas.pack(side="left", fill="both", expand=True)

    table_frame = ttk.Frame(canvas)
    canvas.create_window((0, 0), window=table_frame, anchor="nw")

    # Determine how many ratio sets we have
    num_ratio_sets = len(RATIO_SETS)

    # Each ratio set will have Buy Price, Sell Price, Net Premium columns
    cols_per_ratio = 3

    # Create main header row - Base Strike column
    ttk.Label(table_frame, text="Base Strike", width=10, font=("Arial", 10, "bold"),
              borderwidth=1, relief="solid").grid(row=0, column=0, rowspan=2, padx=1, pady=1, sticky="nsew")

    # Add ratio set headers for this offset
    for i, ratio_set in enumerate(RATIO_SETS):
        ratio_text = f"{ratio_set[0]}x{ratio_set[1]}"
        ratio_col = 1 + (i * cols_per_ratio)
        ttk.Label(table_frame, text=ratio_text, width=24, font=("Arial", 10, "bold"),
                  borderwidth=1, relief="solid").grid(row=0, column=ratio_col,
                                                      columnspan=cols_per_ratio, padx=1, pady=1, sticky="nsew")

        # Add column headers for each ratio set
        ttk.Label(table_frame, text="Buy Price", width=8,
                  borderwidth=1, relief="solid").grid(row=1, column=ratio_col, padx=1, pady=1, sticky="nsew")
        ttk.Label(table_frame, text="Sell Price", width=8,
                  borderwidth=1, relief="solid").grid(row=1, column=ratio_col+1, padx=1, pady=1, sticky="nsew")
        ttk.Label(table_frame, text="Net Premium", width=8,
                  borderwidth=1, relief="solid").grid(row=1, column=ratio_col+2, padx=1, pady=1, sticky="nsew")

    # Get base strikes for this offset
    offset_data = organized_data.get(offset, {})
    base_strikes = sorted(offset_data.keys()) if offset_data else []

    # Fill in the table with data
    for row_idx, base_strike in enumerate(base_strikes):
        # Add the base strike column
        ttk.Label(table_frame, text=str(base_strike), width=10, font=("Arial", 10),
                  borderwidth=1, relief="solid").grid(row=row_idx+2, column=0, padx=1, pady=1, sticky="nsew")

        # Add data for each ratio set
        for i, ratio_set in enumerate(RATIO_SETS):
            ratio = f"{ratio_set[0]}x{ratio_set[1]}"
            ratio_col = 1 + (i * cols_per_ratio)

            # Get data for this ratio or use placeholder
            if ratio in offset_data[base_strike]:
                buy_price = offset_data[base_strike][ratio]["buy_price"]
                sell_price = offset_data[base_strike][ratio]["sell_price"]
                net_premium = offset_data[base_strike][ratio]["net_premium"]

                # Add color coding based on net premium value
                if net_premium < 0:
                    premium_bg = "#ccffcc"  # Light green for negative (favorable for ratio spreads)
                else:
                    premium_bg = "#ffcccc"  # Light red for positive (unfavorable for ratio spreads)
            else:
                buy_price = "-"
                sell_price = "-"
                net_premium = "-"
                premium_bg = "white"

            ttk.Label(table_frame, text=str(buy_price), width=8,
                      borderwidth=1, relief="solid").grid(row=row_idx+2, column=ratio_col,
                                                          padx=1, pady=1, sticky="nsew")
            ttk.Label(table_frame, text=str(sell_price), width=8,
                      borderwidth=1, relief="solid").grid(row=row_idx+2, column=ratio_col+1,
                                                          padx=1, pady=1, sticky="nsew")

            # For net premium, use a regular tk.Label to enable background color
            lbl = tk.Label(table_frame, text=str(net_premium), width=8, bg=premium_bg,
                           borderwidth=1, relief="solid")
            lbl.grid(row=row_idx+2, column=ratio_col+2, padx=1, pady=1, sticky="nsew")

    # Update scrollregion after adding all items
    table_frame.update_idletasks()
    canvas.config(scrollregion=canvas.bbox("all"))

    return table_frame

def create_put_offset_table(notebook, offset, organized_data):
    """Create a table for a specific put offset value"""
    # Create a frame for this offset
    offset_frame = ttk.Frame(notebook)
    notebook.add(offset_frame, text=f"Offset: -{offset}")

    # Create a canvas with scrollbars for the matrix
    canvas = tk.Canvas(offset_frame)
    y_scrollbar = ttk.Scrollbar(offset_frame, orient="vertical", command=canvas.yview)
    x_scrollbar = ttk.Scrollbar(offset_frame, orient="horizontal", command=canvas.xview)

    canvas.configure(yscrollcommand=y_scrollbar.set, xscrollcommand=x_scrollbar.set)
    y_scrollbar.pack(side="right", fill="y")
    x_scrollbar.pack(side="bottom", fill="x")
    canvas.pack(side="left", fill="both", expand=True)

    table_frame = ttk.Frame(canvas)
    canvas.create_window((0, 0), window=table_frame, anchor="nw")

    # Determine how many ratio sets we have
    num_ratio_sets = len(RATIO_SETS)

    # Each ratio set will have Buy Price, Sell Price, Net Premium columns
    cols_per_ratio = 3

    # Create main header row - Base Strike column
    ttk.Label(table_frame, text="Base Strike", width=10, font=("Arial", 10, "bold"),
              borderwidth=1, relief="solid").grid(row=0, column=0, rowspan=2, padx=1, pady=1, sticky="nsew")

    # Add ratio set headers for this offset
    for i, ratio_set in enumerate(RATIO_SETS):
        ratio_text = f"{ratio_set[0]}x{ratio_set[1]}"
        ratio_col = 1 + (i * cols_per_ratio)
        ttk.Label(table_frame, text=ratio_text, width=24, font=("Arial", 10, "bold"),
                  borderwidth=1, relief="solid").grid(row=0, column=ratio_col,
                                                      columnspan=cols_per_ratio, padx=1, pady=1, sticky="nsew")

        # Add column headers for each ratio set
        ttk.Label(table_frame, text="Buy Price", width=8,
                  borderwidth=1, relief="solid").grid(row=1, column=ratio_col, padx=1, pady=1, sticky="nsew")
        ttk.Label(table_frame, text="Sell Price", width=8,
                  borderwidth=1, relief="solid").grid(row=1, column=ratio_col+1, padx=1, pady=1, sticky="nsew")
        ttk.Label(table_frame, text="Net Premium", width=8,
                  borderwidth=1, relief="solid").grid(row=1, column=ratio_col+2, padx=1, pady=1, sticky="nsew")

    # Get base strikes for this offset
    offset_data = organized_data.get(offset, {})
    base_strikes = sorted(offset_data.keys()) if offset_data else []

    # Fill in the table with data
    for row_idx, base_strike in enumerate(base_strikes):
        # Add the base strike column
        ttk.Label(table_frame, text=str(base_strike), width=10, font=("Arial", 10),
                  borderwidth=1, relief="solid").grid(row=row_idx+2, column=0, padx=1, pady=1, sticky="nsew")

        # Add data for each ratio set
        for i, ratio_set in enumerate(RATIO_SETS):
            ratio = f"{ratio_set[0]}x{ratio_set[1]}"
            ratio_col = 1 + (i * cols_per_ratio)

            # Get data for this ratio or use placeholder
            if ratio in offset_data[base_strike]:
                buy_price = offset_data[base_strike][ratio]["buy_price"]
                sell_price = offset_data[base_strike][ratio]["sell_price"]
                net_premium = offset_data[base_strike][ratio]["net_premium"]

                # Add color coding based on net premium value
                if net_premium < 0:
                    premium_bg = "#ccffcc"  # Light green for negative (favorable for ratio spreads)
                else:
                    premium_bg = "#ffcccc"  # Light red for positive (unfavorable for ratio spreads)
            else:
                buy_price = "-"
                sell_price = "-"
                net_premium = "-"
                premium_bg = "white"

            ttk.Label(table_frame, text=str(buy_price), width=8,
                      borderwidth=1, relief="solid").grid(row=row_idx+2, column=ratio_col,
                                                          padx=1, pady=1, sticky="nsew")
            ttk.Label(table_frame, text=str(sell_price), width=8,
                      borderwidth=1, relief="solid").grid(row=row_idx+2, column=ratio_col+1,
                                                          padx=1, pady=1, sticky="nsew")

            # For net premium, use a regular tk.Label to enable background color
            lbl = tk.Label(table_frame, text=str(net_premium), width=8, bg=premium_bg,
                           borderwidth=1, relief="solid")
            lbl.grid(row=row_idx+2, column=ratio_col+2, padx=1, pady=1, sticky="nsew")

    # Update scrollregion after adding all items
    table_frame.update_idletasks()
    canvas.config(scrollregion=canvas.bbox("all"))

    return table_frame

def create_top_premium_table(parent_frame, option_type="calls"):
    """Create a table showing negative net premium results sorted in ascending order"""
    global global_results

    # Get the curve fitting filter state
    check_curve_fit = False
    for widget in parent_frame.winfo_toplevel().winfo_children():
        if hasattr(widget, "curve_fit_var"):
            check_curve_fit = widget.curve_fit_var.get()
            break

    # Create frame for the table
    option_label = "Call" if option_type == "calls" else "Put"
    filter_text = "Curve-Fitted " if check_curve_fit else ""
    table_frame = ttk.LabelFrame(parent_frame, text=f"{filter_text}Negative Net Premium {option_label} Results")
    table_frame.pack(fill="both", expand=True, padx=10, pady=10)

    # Create treeview
    if option_type == "calls":
        columns = ("ratio", "strategy", "base_strike", "higher_strike", "buy_price",
                   "sell_price", "net_premium", "pnl", "percentage_away", "spread_width")
    else:  # puts
        columns = ("ratio", "strategy", "base_strike", "lower_strike", "buy_price",
                   "sell_price", "net_premium", "pnl", "percentage_away", "spread_width")

    tree = ttk.Treeview(table_frame, columns=columns, show="headings", height=10)

    # Define column headings for both call and put tables
    tree.heading("ratio", text="Ratio", command=lambda: sort_table(tree, "ratio", False))
    tree.heading("strategy", text="Strategy", command=lambda: sort_table(tree, "strategy", False))
    tree.heading("base_strike", text="Base Strike", command=lambda: sort_table(tree, "base_strike", False))

    if option_type == "calls":
        tree.heading("higher_strike", text="Higher Strike", command=lambda: sort_table(tree, "higher_strike", False))
    else:  # puts
        tree.heading("lower_strike", text="Lower Strike", command=lambda: sort_table(tree, "lower_strike", False))

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

    if option_type == "calls":
        tree.column("higher_strike", width=80, anchor="center")
    else:  # puts
        tree.column("lower_strike", width=80, anchor="center")

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
    results_list = global_results.get(option_type, [])
    if results_list:
        all_trades = []
        for result in results_list:
            all_trades.extend(result["trades"])

        # Filter only trades with negative net premium
        negative_premium_trades = [trade for trade in all_trades if trade["net_premium"] < 0]

        # Apply curve fitting filter if checked
        if check_curve_fit:
            # We need to filter trades that fit the curve
            # Get the first result for ATM strike
            atm_strike = results_list[0]["atm_strike"]

            # Filter list based on option type
            filtered_trades = []
            for trade in negative_premium_trades:
                if option_type == "calls":
                    base_strike = trade["base_strike"]
                    higher_strike = trade["higher_strike"]
                    # Get option_chain_data from the first result
                    option_chain_data = results_list[0].get("option_chain_data")
                    if option_chain_data and check_call_curve_fit(option_chain_data, atm_strike, base_strike, higher_strike):
                        filtered_trades.append(trade)
                else:  # puts
                    base_strike = trade["base_strike"]
                    lower_strike = trade["lower_strike"]
                    option_chain_data = results_list[0].get("option_chain_data")
                    if option_chain_data and check_put_curve_fit(option_chain_data, atm_strike, base_strike, lower_strike):
                        filtered_trades.append(trade)

        # Sort all negative trades by net premium (ascending - most negative first)
        sorted_negative_trades = sorted(negative_premium_trades, key=lambda x: x["net_premium"])

        # Add to treeview (all negative trades, not just top 3 per ratio)
        for i, trade in enumerate(sorted_negative_trades):
            if option_type == "calls":
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
            else:  # puts
                tree.insert("", "end", values=(
                    f"{trade['buy_qty']}x{trade['sell_qty']}",
                    trade["strategy"],
                    trade["base_strike"],
                    trade["lower_strike"],
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
    """Create the main application UI with a persistent frame for controls"""
    global auto_poll_job

    # Create a persistent frame for controls that won't be destroyed during refresh
    if not hasattr(root, "persistent_frame"):
        persistent_frame = ttk.Frame(root, padding=10)
        persistent_frame.pack(fill="x", pady=5)
        root.persistent_frame = persistent_frame

        # Add note about API tokens
        ttk.Label(persistent_frame, text="Note: Make sure API tokens are valid and expiry dates use YYYY-MM-DD format",
                  font=("Arial", 10, "italic")).pack(anchor="w", pady=5)

        # Add status label
        status_label = ttk.Label(persistent_frame, text="Ready")
        status_label.pack(fill="x", pady=5)
        root.status_label = status_label

        # Create frame for strategy parameters
        strategy_frame = ttk.LabelFrame(persistent_frame, text="Strategy Parameters")
        strategy_frame.pack(fill="x", pady=5, padx=5)

        # ATM Adder
        ttk.Label(strategy_frame, text="ATM Adder:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        root.atm_adder_var = tk.StringVar(value=str(ATM_ADDER))
        ttk.Entry(strategy_frame, textvariable=root.atm_adder_var, width=8).grid(row=0, column=1, padx=5, pady=5, sticky="w")

        # Incremental Adder
        ttk.Label(strategy_frame, text="Incremental Adder:").grid(row=0, column=2, padx=5, pady=5, sticky="w")
        root.incremental_adder_var = tk.StringVar(value=str(INCREMENTAL_ADDER))
        ttk.Entry(strategy_frame, textvariable=root.incremental_adder_var, width=8).grid(row=0, column=3, padx=5, pady=5, sticky="w")

        # Number of Bases
        ttk.Label(strategy_frame, text="Number of Bases:").grid(row=0, column=4, padx=5, pady=5, sticky="w")
        root.num_bases_var = tk.StringVar(value=str(NUM_BASES))
        ttk.Entry(strategy_frame, textvariable=root.num_bases_var, width=8).grid(row=0, column=5, padx=5, pady=5, sticky="w")

        # Offset Values - Using 5 separate text boxes
        ttk.Label(strategy_frame, text="Offset Values:").grid(row=1, column=0, padx=5, pady=5, sticky="w")

        # Create a frame to hold the offset text boxes
        offset_frame = ttk.Frame(strategy_frame)
        offset_frame.grid(row=1, column=1, columnspan=5, padx=5, pady=5, sticky="w")

        # Create 5 text boxes for offset values
        root.offset_vars = []
        for i in range(5):
            offset_var = tk.StringVar()
            # Set initial values from OFFSET_VALUES if available, otherwise empty
            if i < len(OFFSET_VALUES):
                offset_var.set(str(OFFSET_VALUES[i]))

            entry = ttk.Entry(offset_frame, textvariable=offset_var, width=6)
            entry.grid(row=0, column=i, padx=3)
            root.offset_vars.append(offset_var)

            # Add curve fitting checkbox
            # Add curve fitting checkbox
            root.curve_fit_var = tk.BooleanVar(value=False)
            curve_fit_checkbox = ttk.Checkbutton(
                strategy_frame,
                text="Filter for Curve Fitting",
                variable=root.curve_fit_var,
                command=lambda: update_display_filter(root)  # Add this callback
            )
            curve_fit_checkbox.grid(row=2, column=0, columnspan=2, padx=5, pady=5, sticky="w")

            # Add tooltip/help text for curve fitting
            ttk.Label(strategy_frame,
                      text="Curve fitting checks if price ratio decreases as strikes move away from ATM",
                      font=("Arial", 8, "italic")).grid(row=2, column=2, columnspan=4, padx=5, pady=5, sticky="w")

        # Create control frame for buttons and polling options
        control_frame = ttk.Frame(persistent_frame)
        control_frame.pack(fill="x", pady=5)

        # Add fetch button
        fetch_button = ttk.Button(control_frame, text="Fetch & Analyze Data",
                                  command=lambda: fetch_and_analyze(root))
        fetch_button.grid(row=0, column=0, padx=(0, 10))
        root.fetch_button = fetch_button

        # Create auto-poll variable
        root.auto_poll_var = tk.BooleanVar(value=False)

        # Add auto-polling checkbox
        auto_poll_checkbox = ttk.Checkbutton(
            control_frame,
            text="Auto-poll",
            variable=root.auto_poll_var,
            command=lambda: toggle_auto_polling(root)
        )
        auto_poll_checkbox.grid(row=0, column=1, padx=5)

        # Create interval variable
        root.poll_interval_var = tk.StringVar(value="5")

        # Add polling interval input
        ttk.Label(control_frame, text="Interval (sec):").grid(row=0, column=2, padx=5)
        poll_interval_entry = ttk.Entry(control_frame, textvariable=root.poll_interval_var, width=5)
        poll_interval_entry.grid(row=0, column=3, padx=5)
        root.poll_interval_entry = poll_interval_entry

    # Create content frame (this part will be destroyed and rebuilt with new data)
    if hasattr(root, "content_frame"):
        root.content_frame.destroy()

    content_frame = ttk.Frame(root)
    content_frame.pack(fill="both", expand=True, padx=10, pady=5)
    root.content_frame = content_frame

    # Create a notebook for calls vs puts tabs
    option_notebook = ttk.Notebook(content_frame)
    option_notebook.pack(fill="both", expand=True, padx=5, pady=5)

    # Create frames for call and put options
    calls_frame = ttk.Frame(option_notebook)
    puts_frame = ttk.Frame(option_notebook)

    option_notebook.add(calls_frame, text="Call Ratio Spreads")
    option_notebook.add(puts_frame, text="Put Ratio Spreads")

    # Use the consolidated display in each frame
    create_consolidated_display(calls_frame, "calls")
    create_consolidated_display(puts_frame, "puts")

    # Create top premium tables for calls and puts
    create_top_premium_table(calls_frame, "calls")
    create_top_premium_table(puts_frame, "puts")

# Variable to store the auto-polling job ID
auto_poll_job = None

def toggle_auto_polling(root):
    """Toggle auto-polling on or off"""
    global auto_poll_job

    if root.auto_poll_var.get():
        # Auto-polling enabled
        try:
            interval = int(root.poll_interval_var.get())
            if interval < 1:
                messagebox.showerror("Invalid Interval", "Interval must be at least 1 second.")
                root.auto_poll_var.set(False)
                return

            # Disable fetch button when auto-polling
            root.fetch_button.config(state="disabled")

            # Disable the interval entry while polling is active
            root.poll_interval_entry.config(state="disabled")

            # Start auto-polling
            root.status_label.config(text=f"Auto-polling started (every {interval} seconds)")

            # Schedule the first poll
            auto_poll_job = root.after(100, lambda: auto_poll(root, interval))

        except ValueError:
            messagebox.showerror("Invalid Interval", "Please enter a valid number for the polling interval.")
            root.auto_poll_var.set(False)
    else:
        # Auto-polling disabled
        if auto_poll_job is not None:
            root.after_cancel(auto_poll_job)
            auto_poll_job = None

        # Re-enable fetch button
        root.fetch_button.config(state="normal")

        # Re-enable interval entry
        root.poll_interval_entry.config(state="normal")

        root.status_label.config(text="Auto-polling stopped")

def auto_poll(root, interval):
    """Perform auto-polling at the specified interval"""
    global auto_poll_job

    try:
        # Run the fetch and analyze operation
        fetch_and_analyze(root)

        # If auto-polling is still enabled, schedule the next poll
        if root.auto_poll_var.get():
            root.status_label.config(text=f"Auto-polling: Next update in {interval} seconds")
            auto_poll_job = root.after(interval * 1000, lambda: auto_poll(root, interval))
    except Exception as e:
        # If there's an error, stop auto-polling and show error
        root.auto_poll_var.set(False)
        root.fetch_button.config(state="normal")
        root.poll_interval_entry.config(state="normal")

        # Cancel any pending auto-poll job
        if auto_poll_job is not None:
            root.after_cancel(auto_poll_job)
            auto_poll_job = None

        root.status_label.config(text=f"Auto-polling error: {str(e)}")
        print(f"Auto-polling error: {str(e)}")
        import traceback
        traceback.print_exc()

def fetch_and_analyze(root):
    """Fetch and analyze data for all instruments"""
    global global_results

    # Clear global_results to start fresh
    global_results = {
        "calls": [],
        "puts": []
    }

    root.status_label.config(text="Fetching and analyzing data... Please wait")
    root.update()

    try:
        # Get strategy parameters from UI
        try:
            atm_adder = int(root.atm_adder_var.get())
            incremental_adder = int(root.incremental_adder_var.get())
            num_bases = int(root.num_bases_var.get())

            check_curve_fit = root.curve_fit_var.get()
            print(f"Curve fitting checkbox state: {check_curve_fit}")

            # Get offset values from the five text boxes, filter out empty ones
            offset_values = []
            for offset_var in root.offset_vars:
                value = offset_var.get().strip()
                if value:  # Only add non-empty values
                    try:
                        offset_values.append(int(value))
                    except ValueError:
                        # Skip invalid entries
                        continue

            # Validate inputs
            if atm_adder < 0 or incremental_adder < 0 or num_bases <= 0 or not offset_values:
                raise ValueError("Invalid parameter values")

            strategy_params = {
                'atm_adder': atm_adder,
                'incremental_adder': incremental_adder,
                'num_bases': num_bases,
                'offset_values': offset_values
            }

            print(f"Using strategy parameters: {strategy_params}")

        except (ValueError, AttributeError) as e:
            messagebox.showerror("Invalid Parameters", f"Please enter valid numbers for all strategy parameters: {str(e)}")
            root.status_label.config(text="Ready")
            return

        # Process all instruments
        for instrument in INSTRUMENTS:
            instrument_key = instrument["instrument_key"]
            expiry_date = instrument["expiry_date"]
            symbol_name = instrument["symbol_name"]
            lot_size = instrument["lot_size"]

            root.status_label.config(text=f"Fetching data for {symbol_name}...")
            root.update()

            # Fetch option chain data
            option_chain = fetch_option_chain(instrument_key, expiry_date)

            if option_chain:
                root.status_label.config(text=f"Analyzing {symbol_name} call options...")
                root.update()

                # Analyze call option chain data with strategy parameters from UI
                analyze_call_option_chain(option_chain, symbol_name, lot_size, strategy_params, check_curve_fit)

                root.status_label.config(text=f"Analyzing {symbol_name} put options...")
                root.update()

                # Analyze put option chain data with strategy parameters from UI
                analyze_put_option_chain(option_chain, symbol_name, lot_size, strategy_params, check_curve_fit)
            else:
                root.status_label.config(text=f"Failed to fetch data for {symbol_name}")
                root.update()
                time.sleep(1)

        # Only rebuild the content part of the UI, leaving the control panel intact
        if hasattr(root, "content_frame"):
            root.content_frame.destroy()

        content_frame = ttk.Frame(root)
        content_frame.pack(fill="both", expand=True, padx=10, pady=5)
        root.content_frame = content_frame

        # Create a notebook for calls vs puts tabs
        option_notebook = ttk.Notebook(content_frame)
        option_notebook.pack(fill="both", expand=True, padx=5, pady=5)

        # Create frames for call and put options
        calls_frame = ttk.Frame(option_notebook)
        puts_frame = ttk.Frame(option_notebook)

        option_notebook.add(calls_frame, text="Call Ratio Spreads")
        option_notebook.add(puts_frame, text="Put Ratio Spreads")

        # Use the consolidated display in each frame
        create_consolidated_display(calls_frame, "calls")
        create_consolidated_display(puts_frame, "puts")

        # Create top premium tables for calls and puts
        create_top_premium_table(calls_frame, "calls")
        create_top_premium_table(puts_frame, "puts")

        if check_curve_fit:
            root.status_label.config(text="Analysis complete - Results filtered for curve fitting")
        else:
            root.status_label.config(text="Analysis complete - All results shown")

    except Exception as e:
        root.status_label.config(text=f"Error: {str(e)}")
        print(f"Error fetching/analyzing data: {str(e)}")
        import traceback
        traceback.print_exc()

def update_display_filter(root):
    """Update the display based on the curve fitting filter state without re-fetching data"""
    # Only update if we already have data
    if not global_results["calls"] and not global_results["puts"]:
        return

    # Rebuild just the content part of the UI
    if hasattr(root, "content_frame"):
        root.content_frame.destroy()

    content_frame = ttk.Frame(root)
    content_frame.pack(fill="both", expand=True, padx=10, pady=5)
    root.content_frame = content_frame

    # Create a notebook for calls vs puts tabs
    option_notebook = ttk.Notebook(content_frame)
    option_notebook.pack(fill="both", expand=True, padx=5, pady=5)

    # Create frames for call and put options
    calls_frame = ttk.Frame(option_notebook)
    puts_frame = ttk.Frame(option_notebook)

    option_notebook.add(calls_frame, text="Call Ratio Spreads")
    option_notebook.add(puts_frame, text="Put Ratio Spreads")

    # Use the consolidated display in each frame - now passing the filter state
    create_consolidated_display(calls_frame, "calls")
    create_consolidated_display(puts_frame, "puts")

    # Create top premium tables for calls and puts - also passing filter state
    create_top_premium_table(calls_frame, "calls")
    create_top_premium_table(puts_frame, "puts")

    # Update status to show filtering state
    if root.curve_fit_var.get():
        root.status_label.config(text="Displaying results filtered for curve fitting")
    else:
        root.status_label.config(text="Displaying all results")

def main():
    """Main function"""
    global auto_poll_job

    root = tk.Tk()
    root.title("Option Ratio Spread Analysis")  # Updated title to reflect both call and put analysis
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

    # Clean up auto polling job when window is closed
    def on_closing():
        if auto_poll_job is not None:
            root.after_cancel(auto_poll_job)
        root.destroy()

    root.protocol("WM_DELETE_WINDOW", on_closing)

    # Start the main loop
    root.mainloop()

if __name__ == "__main__":
    main()