# Fuzzy Matching Application

Usage Care Brand Fuzzy Matching

# Situation

The team I'm working with was using Excel's Fuzzy Lookup tool to extract brand names from free string values. I tried to find an alternative that easier for team to use and outperform hte Fuzzy Lookup. 

# Approach

Textdistance library was working great with finding similarities for longer string values based on my experience, but didn't work out well with shorter text values. This Python script applies the "token_sort_ratio" methods from the fuzzywuzzy library. We have seen that this method outperforms MS Excel's Fuzzy Lookup for shorter text (one or two words).

# Arguments

**fuzzy_base**: the file location, and the sheet name includes the column to use as the fuzzy matching base.

**brand_base**: the file location, and the sheet name includes the column to use as the fuzzy matching lookup.

**base_column**: column name of fuzzy matching base

**brand_column**: column name of fuzzy matching lookup

**cutoff_value**: The similarity score between 0 and 100. Based on our experience 75/80 is optimum as a cutoff point.

# Attention!

Please ensure that you adjusted the input (Excel) and output (CSV) in the script for your personal needs before running the job. 

# Author:
Adnan Yoztyurk
