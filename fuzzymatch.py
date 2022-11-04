# -*- coding: utf-8 -*-
"""Fuzzy Matching Applicaton

Description:
    This Python script applies the "token_sort_ratio" methods from the fuzzywuzzy library. We have seen that this method
    outperforms MS Excel's Fuzzy Lookup for shorter text (one or two words).

Requirements:
    numpy==1.23.1
    fuzzywuzzy==0.18.0
    pandas==1.4.1

Arguments:
    fuzzy_base: the file location, and the sheet name includes the column to use as the fuzzy matching base.
    brand_base: the file location, and the sheet name includes the column to use as the fuzzy matching lookup.
    base_column: column name of fuzzy matching base
    brand_column: column name of fuzzy matching lookup
    cutoff_value: The similarity score between 0 and 100. Based on our experience 75/80 is optimum as a cutoff point.

Attention! Please ensure that you adjusted the input (Excel) and output (CSV) in the script for your personal needs
before running the job.

Author: Adnan Yoztyurk
"""

# Import required libraries
from fuzzywuzzy import fuzz
from fuzzywuzzy import process
import pandas as pd
import numpy as np


# Defined function uses fuzzywuzzy processor
def fuzzy_match(x, choices, scorer, cutoff):
    return process.extractOne(
        x, choices=choices, scorer=scorer, score_cutoff=cutoff
    )


# Defined function to feed the inputs to apply fuzzy matching
def fuzzy_match_apply(func, cut, base, base_col, lookup, lookup_col, new_col):
    base.loc[:, new_col] = base.loc[:, base_col].apply(
        fuzzy_match,
        args=(
            lookup.loc[:, lookup_col],
            func,
            cut
        )
    )


# Please change the file path and page name and make sure they are in the correct location.
fuzzy_base = pd.read_excel('Please enter the location of input Excel file including the base column.')

# Please change the file path and page name and make sure they are in the correct location.
brand_base = pd.read_excel('Please enter the location of input Excel file including the reference column.')
brand_base['index'] = brand_base.index

# Please be ensure that you checked the column names are correct below
base_column = 'Descr'
brand_column = 'Brand'
cutoff_value = 75

fuzzy_base[base_column] = fuzzy_base[base_column].fillna('NullValue').apply(str)

# Apply token_sort_ratio method
fuzzy_match_apply(fuzz.token_sort_ratio, cutoff_value, fuzzy_base, base_column, brand_base, brand_column, 'Sort')

# Data transformations
fuzzy_base.replace(to_replace=[None], value=np.nan, inplace=True)
fuzzy_base_sort = fuzzy_base['Sort'].apply(pd.Series)
fuzzy_base_sort.columns = ['Sort_desc', 'Sort_score', 'Sort_index']
fuzzy_base_sort_merge = pd.merge(fuzzy_base_sort, brand_base, how="left", left_on="Sort_index", right_on="index")
fuzzy_base_sort_merge.rename(columns={'Ans_Code': 'Sort_ans_code', 'Brand': 'Sort_brand'}, inplace=True)
fuzzy_base_interim = pd.concat([fuzzy_base, fuzzy_base_sort_merge], axis=1)
fuzzy_base_final = fuzzy_base_interim[[base_column, 'Sort_desc', 'Sort_ans_code', 'Sort_score']]

# Please change the file path and make sure they are in the correct location for the output CSV file.
fuzzy_output = 'Please enter the location of output CSV file.'
fuzzy_base_final.to_csv(fuzzy_output)

print('Task completed successfully!')

# End of the script
