import pandas as pd

# Create a dataframe to hold the data being read
marriage = pd.read_excel('part2/marriage_rate_input.xlsx',
                       #skiprows=5,                      # skip 1st 5 rows of excel file
                       header=[0, 1],                   # header is now in the first and 2nd row
                       index_col=[0]                    # index only includes 1st column
                        )

# print(marriage.columns.values)    # test column output
# print(marriage.head())          # test output
marriage = marriage.stack()       # pivot years with Year column in multi-index
marriage = marriage.reset_index() # reset the index to fill all the missing values
col3 = marriage.columns[2]        # set column 3 as marriage rate

marriage.rename(columns={marriage.columns[0] : 'State',      # rename first two columns
                         marriage.columns[1] : 'Year',
                         marriage.columns[2] : col3})

# print(marriage.head())           # test output

marriage.to_excel(excel_writer='part2/marriage_rate_output.xls',  # write to the given file
                  sheet_name='marriage_rate',  # name the sheet
                  index=False)                 # do not include the pandas index