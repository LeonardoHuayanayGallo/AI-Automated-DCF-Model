# Investment Banking AI Assistant: Automated DCF Valuation & LLM-Driven 10-K Analysis

![Python](https://img.shields.io/badge/Python-3.10+-blue)
![License](https://img.shields.io/badge/License-MIT-green)
![SEC EDGAR](https://img.shields.io/badge/Data-SEC%20EDGAR-orange)
![HuggingFace](https://img.shields.io/badge/LLM-Llama%203-yellow)

## Overview

An end-to-end AI-powered investment banking tool that automates the most 
time-intensive parts of fundamental equity analysis:

1. **Retrieves** the most recent 10-K filing directly from the SEC EDGAR API
2. **Extracts** the Management's Discussion & Analysis (MD&A) section
3. **Parses** qualitative forward-looking language using Meta Llama 3 (open-source LLM)
   to generate structured financial assumptions
4. **Builds** a fully formatted 5-year DCF valuation model in Excel, complete 
   with a two-dimensional sensitivity table (WACC × Terminal Growth Rate)

This replicates a core workflow in investment banking due diligence — turning 
qualitative management commentary into quantitative model inputs — in a 
fully automated pipeline.


![Image Alt](https://github.com/LeonardoHuayanayGallo/AI-Automated-DCF-Model/blob/aafa465c4185085d40b05e841f33ce9f59bde316/DCF%20Model.png)

![Image Alt](https://github.com/LeonardoHuayanayGallo/AI-Automated-DCF-Model/blob/aafa465c4185085d40b05e841f33ce9f59bde316/Sensitivity%20Analysis.png)
