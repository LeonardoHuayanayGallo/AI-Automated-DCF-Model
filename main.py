"""
main.py
Full pipeline: SEC 10-K → LLM extraction → DCF model → Excel output
Usage: python main.py --ticker AAPL --revenue 383285
"""

import argparse
import os
from dotenv import load_dotenv
import sec_extractor
import llm_parser
import dcf_model

load_dotenv()

def main():
    parser = argparse.ArgumentParser(description="AI-Powered DCF Valuation Engine")
    parser.add_argument("--ticker",  required=True, help="Stock ticker (e.g. AAPL)")
    parser.add_argument("--revenue", required=True, type=float,
                        help="Most recent annual revenue in $M (from 10-K)")
    parser.add_argument("--wacc",    type=float, default=0.10, help="WACC (default: 10%)")
    parser.add_argument("--tgr",     type=float, default=0.025,
                        help="Terminal growth rate (default: 2.5%)")
    args = parser.parse_args()

    # Step 1: Extract MD&A from SEC EDGAR
    mda_text = sec_extractor.run(args.ticker)

    # Step 2: LLM extracts financial assumptions
    assumptions = llm_parser.extract_assumptions(mda_text)

    # Step 3: Build DCF model and export to Excel
    output_path = dcf_model.build_dcf(
        ticker       = args.ticker,
        base_revenue = args.revenue,
        assumptions  = assumptions,
        wacc         = args.wacc,
        terminal_growth = args.tgr
    )

    print(f"\n✅ DCF model complete. File: {output_path}")

if __name__ == "__main__":
    main()