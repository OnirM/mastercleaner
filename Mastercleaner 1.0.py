import pandas as pd
def remove_non_duplicate_rows(file_path):
  # Load the Excel file in a DataFrame
  df = pd.read_excel(file_path, header=0)
  # Remove the rows with no duplicates in the column 'Column1'
  df = df[df.duplicated(subset='Column1', keep=False)]
  # Formatt the 'Column2' column in date formato yyy/mm/gg
  df['Column2'] = pd.to_datetime(df['Column2'], dayfirst=True)
  return df

def clean_duplicate(df):
  # From the original DataFrame, creates a new DataFrame that contains only the rows with just one duplicate (saves also the duplicate)
  duplicate_simple_df = df[df['Column1'].isin(df['Column1'].value_counts()[df['Column1'].value_counts() == 2].index)]
  
  # In the duplicate_simple DataFrame, deletes the rows that have both the 'Column1' and the 'Column3' columns duplicated and keeps the rest
  duplicate_simple_df = duplicate_simple_df[~duplicate_simple_df.duplicated(subset=['Column1', 'Column3'], keep=False)]

  # From the original DataFrame, creates a new DataFrame that contains only the rows with more than one duplicate (saves also the duplicates)
  duplicate_multiple_df = df[df['Column1'].isin(df['Column1'].value_counts()[df['Column1'].value_counts() > 2].index)]

  # In the duplicate_multiple DataFrame, sorts the DataFrame by value of the column 'Column2', identifies the rows with both the 'Column1' and 'Column3' column duplicated and keeps only the one with the highest value in 'Column2' (keeps also the duplicate with a different value in the 'Column3' column)
  duplicate_multiple_df = (
    duplicate_multiple_df.sort_values(by='Column2')
    .drop_duplicates(subset=['Column1', 'Column3'], keep='last')
  )

  # In the duplicate_multiple DataFrame, after the first function, remove the rows with no duplicates in the 'Column1' row.
  duplicate_multiple_df = duplicate_multiple_df[duplicate_multiple_df.duplicated(subset='Column1', keep=False)]  
  
  # Merges the two DataFrames
  df_unified = pd.concat([duplicate_simple_df, duplicate_multiple_df])

  # Asks the user for a file name
  file_name = input("Insert the name of your new clean excel file: ")
  
  # Saves the DataFrame in Excel with the prompted file name
  df_unified.to_excel('C:\Path\To\Folder' + '/' + file_name + '.xlsx', index=False)

def main():
  # Ask the user for the file path
  file_path = input("Insert the path to your dirty excel file: ")
  df = remove_non_duplicate_rows(file_path)
  clean_duplicate(df)

if __name__ == "__main__":
  main()