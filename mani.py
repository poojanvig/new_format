import pandas as pd

data = {
    'Day': [1, 2, 3, 4, 5],
    'Sep-2020': [1, 2, 3, 4, 5],
    'Oct-2020': [6, 7, 8, 9, 10],
    'Nov-2020': [11, 12, 13, 14, 15],
    'Dec-2020': [16, 17, 18, 19, 20],
    'Jan-2021': [21, 22, 23, 24, 25],
    'Feb-2021': [26, 27, 28, 29, 30],
    'Mar-2021': [31, 32, 33, 34, 35],

}


def avgs_df(df):
    # quarterly_avg
    if df.shape[1] > 3:
        df_chi_list_1 = []
        # Iterate through every three columns in the original DataFrame
        for i in range(1, df.shape[1], 3):
            # Get the current three columns
            subset = df.iloc[:, i:i + 3]
            if subset.shape[1] < 3:
                new_row = 0.0
            else:
                new_row = subset.iloc[-2].sum() / 3
            subset.loc[len(subset)] = new_row
            df_chi_list_1.append(subset)
        result_df = pd.concat(df_chi_list_1, axis=1)
        new_row = pd.Series({'Day': 'Quarterly Average'})
        df = df._append(new_row, ignore_index=True)
        result_df.insert(0, 'Day', df['Day'])
        df = result_df

        # half - yearly avg
        if df.shape[1] > 6:
            df_chi_list_2 = []
            # Iterate through every three columns in the original DataFrame
            for i in range(1, df.shape[1], 6):
                # Get the current three columns
                subset = df.iloc[:, i:i + 6]
                if subset.shape[1] < 6:
                    new_row = 0.0
                else:
                    new_row = subset.iloc[-3].sum() / 6
                subset.loc[len(subset)] = new_row
                df_chi_list_2.append(subset)
            result_df = pd.concat(df_chi_list_2, axis=1)
            new_row = pd.Series({'Day': 'Half-Yearly Average'})
            df = df._append(new_row, ignore_index=True)
            result_df.insert(0, 'Day', df['Day'])
            df = result_df

            # yearly avg
            if df.shape[1] > 12:
                df_chi_list_3 = []
                # Iterate through every three columns in the original DataFrame
                for i in range(1, df.shape[1], 12):
                    # Get the current three columns
                    subset = df.iloc[:, i:i + 12]
                    if subset.shape[1] < 12:
                        new_row = 0.0
                    else:
                        new_row = subset.iloc[-4].sum() / 12
                    subset.loc[len(subset)] = new_row
                    df_chi_list_3.append(subset)
                result_df = pd.concat(df_chi_list_3, axis=1)
                new_row = pd.Series({'Day': 'Yearly Average'})
                df = df._append(new_row, ignore_index=True)
                result_df.insert(0, 'Day', df['Day'])
                df = result_df


            else:
                new_row = pd.Series({'Day': 'Yearly Average'})
                df = df._append(new_row, ignore_index=True)

        else:
            new_row = pd.Series({'Day': 'Half-Yearly Average'})
            df = df._append(new_row, ignore_index=True)

    else:
        new_row = pd.Series({'Day': 'Quarterly Average'})
        df = df._append(new_row, ignore_index=True)

    return df


df = pd.DataFrame(data)
eod_df = avgs_df(df)
