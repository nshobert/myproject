import os
import pandas as pd
import numpy as np
import tabula

# declare boring designations
borings = ['MB-01',
           'MB-02',
           'MB-02A',
          ]

# declare pile driving hammers
hammers = ['MENCK MHU 500T',
           'MENCK MHU 800S',
           'PILECO D180-32',
           'PILECO D225-22',
          ]

# loop through directories and save WEAP PDF output to Excel workbooks
for boring in borings:
    for hammer in hammers:
        # declare WEAP directory
        fp = os.getcwd() + r'\%s\Driveability\%s' %(boring, hammer)
        
        # get file names in directory
        fns = os.listdir(fp)
        
        # get PDF file names in directory
        pdfs = [fn for fn in fns if fn[-4:]=='.pdf']
        
        # loop through WEAP PDF output files
        for pdf in pdfs:
            # read 2nd sheet of WEAP PDF output files
            df = tabula.read_pdf(os.path.join(fp, pdf), pages=2)[0]

            # tabula joins all columns after third due to first line on 2nd page of WEAP PDF output file
            # clean up header names and rows using blank space as a delimeter
            new_header = df.iloc[1]
            df = df[2:]
            df.columns = new_header

            # setup column names and split third column, then join
            cols = ['Rshaft', 'Rtoe', 'Blow Ct', 'Mx C-Str.', 'Mx T-Str.', 'Stroke', 'ENTHRU', 'Hammer']
            new_df = df[['Depth', 'Rut']].copy()

            # if MENCK hammer, join last two columns as they represent the hammer name, then drop last column
            _temp_df = df.iloc[:, 2].str.split(' ', expand=True)
            
            if 'MENCK' in hammer:
                _temp_df[7] = _temp_df[7] + ' ' + _temp_df[8]
                _temp_df = _temp_df.iloc[:, :-1]

            new_df[cols] = _temp_df

            # if MENCK hammer, replace nan value at Hammer column first row
            if 'MENCK' in hammer:
                if np.isnan(new_df.iloc[0]['Hammer'])==True:
                    new_df.iloc[0, new_df.columns.get_loc('Hammer')] = '-'

            # first row of data are the units for the column headers
            # join the units in the column headers and re-create dataframe
            cols = ['%s (%s)' %(header, unit) for header, unit in zip(new_df.columns, new_df.iloc[0])]
            new_df.columns = cols
            new_df = new_df[1:].reset_index(drop=True)

            # convert all but last column values to float
            cols = new_df.columns
            new_df[cols[:-1]] = new_df[cols[:-1]].apply(pd.to_numeric, errors='coerce')


            # read 3rd sheet of WEAP PDF output files
            df = tabula.read_pdf(os.path.join(fp, pdf), pages=3)[0]
            
            # if MENCK hammer, tabula joins last two columns for some reason...
            # split these columns and rename headers accordingly
            if 'MENCK' in hammer:
                _temp_df = df.iloc[:, 8].str.split(' ', expand=True)
                _temp_df[1] = _temp_df[1] + ' ' + _temp_df[2]
                _temp_df = _temp_df.iloc[:, :-1]

                # join dataframes
                df = df.iloc[:, :-1].copy()
                df = pd.merge(df, _temp_df, left_index=True, right_index=True)
            
            df.columns = cols

            # convert all but last column values to float
            cols = df.columns
            df[cols[:-1]] = df[cols[:-1]].apply(pd.to_numeric, errors='coerce')


            # combine 2nd and 3rd sheet of PDF WEAP output file dataframes
            new_df = pd.concat([new_df, df]).reset_index(drop=True)
            
            # export dataframe to Excel workbook
            out_fp = os.path.join(fp, pdf.replace('.pdf', '.xlsx'))
            new_df.to_excel(out_fp, sheet_name='Summary')  