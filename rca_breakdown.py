import os
import sys
import pandas as pd

# If you want to hardcode the output file path, set this to a valid absolute path (including .xlsx).
# Otherwise leave as None to be prompted at runtime.
DEFAULT_OUTPUT_FILE = './output/rca_breakdown.xlsx'

def identify_mistaken_policies(row):
    correct_policies = eval(
        row['Correct Email Policy List']
        if row['Method'] == 'Call'
        else row['Correct Call Policy List']
    )
    sim_policies = eval(row['Policy List'])
    correct_set = set(correct_policies)
    sim_set = set(sim_policies)
    fp = list(sim_set - correct_set)
    fn = list(correct_set - sim_set)
    return fp, fn

def process_data(input_file, output_file):
    try:
        df = pd.read_excel(input_file, sheet_name='Sheet1')
    except Exception as e:
        print(f"Error reading '{input_file}': {e}", file=sys.stderr)
        sys.exit(1)

    df[['FP', 'FN']] = df.apply(identify_mistaken_policies, axis=1, result_type='expand')
    mistakes_df = df[df['Is_Task_Correct'] == 'wrong']

    controllable = ['Yes', 'Both Wrong']
    non_controllable = ['No', 'Both Right']

    # MARKET LEVEL FP / FN
    market_fp = (mistakes_df
                 .explode('FP').dropna(subset=['FP'])
                 .groupby(['Market','FP'])
                 .size().reset_index(name='FP Count')
                 .rename(columns={'FP':'Policy'}))

    market_fn = (mistakes_df
                 .explode('FN').dropna(subset=['FN'])
                 .groupby(['Market','FN'])
                 .size().reset_index(name='FN Count')
                 .rename(columns={'FN':'Policy'}))

    market_level = pd.merge(market_fp, market_fn, on=['Market','Policy'], how='outer').fillna(0)
    market_level['Mistakes'] = market_level['FP Count'] + market_level['FN Count']
    market_level[['FP Count','FN Count','Mistakes']] = market_level[['FP Count','FN Count','Mistakes']].astype(int)

    # CONTROLLABLE vs NON-CONTROLLABLE helper
    def group_counts(df, subset_vals, col_name, out_name):
        return (df[df['Controllable'].isin(subset_vals)]
                .explode(col_name).dropna(subset=[col_name])
                .groupby(['Market', col_name])
                .size().reset_index(name=out_name)
                .rename(columns={col_name: 'Policy'}))

    ctrl_fp = group_counts(mistakes_df, controllable, 'FP', 'Controllable FP Count')
    ctrl_fn = group_counts(mistakes_df, controllable, 'FN', 'Controllable FN Count')
    nonctrl_fp = group_counts(mistakes_df, non_controllable, 'FP', 'Non-Controllable FP Count')
    nonctrl_fn = group_counts(mistakes_df, non_controllable, 'FN', 'Non-Controllable FN Count')

    ctrl = pd.merge(ctrl_fp, ctrl_fn, on=['Market','Policy'], how='outer').fillna(0)
    ctrl['Controllable Count'] = (ctrl['Controllable FP Count'] + ctrl['Controllable FN Count']).astype(int)

    nonctrl = pd.merge(nonctrl_fp, nonctrl_fn, on=['Market','Policy'], how='outer').fillna(0)
    nonctrl['Non-Controllable Count'] = (nonctrl['Non-Controllable FP Count'] + nonctrl['Non-Controllable FN Count']).astype(int)

    market_level = (market_level
                    .merge(ctrl[['Market','Policy','Controllable Count']], on=['Market','Policy'], how='left')
                    .merge(nonctrl[['Market','Policy','Non-Controllable Count']], on=['Market','Policy'], how='left')
                    .fillna(0))
    market_level[['Controllable Count','Non-Controllable Count']] = market_level[['Controllable Count','Non-Controllable Count']].astype(int)

    # Calculate impacts
    totals = df.groupby('Market').size().reset_index(name='Total Cases')
    market_level = market_level.merge(totals, on='Market')
    market_level['Impact Controllable'] = (market_level['Controllable Count'] / market_level['Total Cases'] * 100).round(2).astype(str) + '%'
    market_level['Impact Non-Controllable'] = (market_level['Non-Controllable Count'] / market_level['Total Cases'] * 100).round(2).astype(str) + '%'
    market_level['Overall Impact'] = (market_level['Mistakes'] / market_level['Total Cases'] * 100).round(2).astype(str) + '%'
    market_level.drop(columns=['Total Cases'], inplace=True)

    market_level = market_level[
        ['Market','Policy','Mistakes','FP Count','FN Count','Overall Impact',
         'Impact Controllable','Impact Non-Controllable','Controllable Count','Non-Controllable Count']
    ].sort_values(by='Mistakes', ascending=False)

    # --- Added RCA count columns for each mistaken policy ---
    rca_fp = (mistakes_df.explode('FP').dropna(subset=['FP'])
              .groupby(['Market','FP','RCA']).size().reset_index(name='count')
              .rename(columns={'FP':'Policy'}))
    rca_fn = (mistakes_df.explode('FN').dropna(subset=['FN'])
              .groupby(['Market','FN','RCA']).size().reset_index(name='count')
              .rename(columns={'FN':'Policy'}))
    rca_all = pd.concat([rca_fp, rca_fn], axis=0)
    rca_pivot = (rca_all.groupby(['Market','Policy','RCA'])['count']
                 .sum().unstack(fill_value=0).reset_index())
    market_level = market_level.merge(rca_pivot, on=['Market','Policy'], how='left').fillna(0)
    # Convert RCA columns to integer type
    rca_cols = [c for c in market_level.columns if c not in
                ['Market','Policy','Mistakes','FP Count','FN Count','Overall Impact',
                 'Impact Controllable','Impact Non-Controllable','Controllable Count','Non-Controllable Count']]
    market_level[rca_cols] = market_level[rca_cols].astype(int)
    # --- End Added RCA count columns ---

    # Write only market sheets
    try:
        with pd.ExcelWriter(output_file) as writer:
            for m in market_level['Market'].unique():
                market_level[market_level['Market']==m].to_excel(writer, sheet_name=f'Market_{m}', index=False)
    except Exception as e:
        print(f"Error writing '{output_file}': {e}", file=sys.stderr)
        sys.exit(1)

    print(f"Processing complete. Output saved to: {output_file}")


def main():
    # get input file
    input_file = input("Enter path to the input Excel file: ").strip()
    if not os.path.isfile(input_file):
        print(f"Input file '{input_file}' not found.", file=sys.stderr)
        sys.exit(1)

    # determine output file
    if DEFAULT_OUTPUT_FILE:
        output_file = DEFAULT_OUTPUT_FILE
        print(f"Using hardcoded output file: {output_file}")
    else:
        output_file = input("Enter desired output file path (including .xlsx): ").strip()

    process_data(input_file, output_file)

if __name__ == "__main__":
    main()