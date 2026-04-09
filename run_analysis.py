"""CLI entry — same output as Analysis.txt (print + Excel on disk)."""
from market_analysis import (
    DEFAULT_TICKERS,
    DEFAULT_TOP_N,
    DEFAULT_VOLUME_SPIKE_THRESHOLD,
    default_excel_filename,
    fetch_stock_data,
    get_market_movers,
    save_results_to_excel,
)

if __name__ == "__main__":
    print("Fetching stock data...")
    df = fetch_stock_data(DEFAULT_TICKERS)

    if df.empty:
        print("No valid stock data found.")
    else:
        most_active, top_gainers, top_losers, volume_spikes = get_market_movers(
            df, top_n=DEFAULT_TOP_N, spike_threshold=DEFAULT_VOLUME_SPIKE_THRESHOLD
        )

        print("\n=== MOST ACTIVE STOCKS ===")
        print(most_active.to_string(index=False))

        print("\n=== TOP GAINERS ===")
        print(top_gainers.to_string(index=False))

        print("\n=== TOP LOSERS ===")
        print(top_losers.to_string(index=False))

        print("\n=== VOLUME SPIKES ===")
        if not volume_spikes.empty:
            print(volume_spikes.to_string(index=False))
        else:
            print("No stocks met the volume spike threshold.")

        output_file = default_excel_filename()
        save_results_to_excel(most_active, top_gainers, top_losers, volume_spikes, output_file)
        print(f"\nSaved results to: {output_file}")
