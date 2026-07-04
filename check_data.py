import pandas as pd

xlsx_path = 'source/bearings.xlsx'

try:
    df = pd.read_excel(xlsx_path)
    
    print("DataFrame shape:", df.shape)
    print("\nFirst 5 rows:")
    print(df.iloc[:5, :7])
    
    print("\nColumn names:")
    for idx, col in enumerate(df.columns):
        print(f"  {idx}: {col}")
    
    # Columns: 0=_1, 1=_2, 2=_3, 3=d, 4=D, 5=B, ... 12=category
    
    print("\nRows where d=3 AND D=10:")
    result = df[(df.iloc[:, 3] == 3) & (df.iloc[:, 4] == 10)]
    print(f"Count: {len(result)}")
    if len(result) > 0:
        print(result[['_1', '_2', '_3']].values)
    
    print("\nRows where d=3 AND D=13:")
    result = df[(df.iloc[:, 3] == 3) & (df.iloc[:, 4] == 13)]
    print(f"Count: {len(result)}")
    if len(result) > 0:
        print(result[['_1', '_2', '_3']].values)
    
    print("\nRows where d=3 AND D in (10,13):")
    result = df[(df.iloc[:, 3] == 3) & (df.iloc[:, 4].isin([10, 13]))]
    print(f"Count: {len(result)}")
    if len(result) > 0:
        for idx, row in result.iterrows():
            print(f"  {row.iloc[0]} - {row.iloc[1]} - d:{row.iloc[3]} D:{row.iloc[4]} B:{row.iloc[5]}")
    
except Exception as e:
    print(f"Error: {type(e).__name__}: {e}")
    import traceback
    traceback.print_exc()
