import pandas as pd

class Allowances:
    def __init__(self, effectiveDate=None, filename=None, verbose=False):
        self.data = None			# a dataframe containing information loaded from a file and cleaned
        self.effectiveDate = effectiveDate if effectiveDate is not None else pd.to_datetime('today').strftime('%Y-%m-%d')
        inputFilename = filename if filename is not None else 'data/Allowances.csv'

        if verbose:
            print(f'Loading Allowances data from {inputFilename} for {self.effectiveDate}')

        df = pd.read_csv(inputFilename)
        df['EffectiveDate'] = pd.to_datetime(df['EffectiveDate'])

        # convert the rates to a percentage
        df['PostingRate'] = df['PostingRate'] * 0.01
        df['HazardRate'] = df['HazardRate'] * 0.01
        df = df.groupby('PostName').apply(lambda x: x.loc[x['EffectiveDate'] <= self.effectiveDate].sort_values(by='EffectiveDate', ascending=False).head(1)).reset_index(drop=True)
        self.data = df

if __name__ == '__main__':
    import sys

    inputFilename = sys.argv[1] if len(sys.argv) > 1 else None
    effectiveDate = pd.to_datetime('2023-06-01').strftime('%Y-%m-%d')   

    allowances = Allowances(effectiveDate=effectiveDate)

    print('Allowances data:')
    print(allowances.data)