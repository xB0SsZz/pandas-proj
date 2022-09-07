import pandas as pd
import numpy as np
import xlsxwriter

EXCEL_NAME = 'example.xlsx'




def main():
    df = pd.read_excel(EXCEL_NAME)

    output_files = list()

    # removing row 1:
    df.to_excel(EXCEL_NAME.partition('.')[0] + "_croped.xlsx", index=False, header=False)
    croped_df = pd.read_excel(EXCEL_NAME.partition('.')[0] + "_croped.xlsx")

    # removing column W:
    croped_df.drop(croped_df.columns[22], axis=1, inplace=True)
    croped_df.to_excel(EXCEL_NAME.partition('.')[0] + "_croped.xlsx", index=False, header=True)
    print(croped_df)

    header = croped_df.columns.values.tolist()

    writer_croped = pd.ExcelWriter(EXCEL_NAME.partition('.')[0] + "_croped.xlsx", engine='xlsxwriter')
    

    rows_to_delete = [-1]
    # reading data_config_file.txt:
    with open('data_config_file.txt', 'r') as f:
        for line in f:
            output_file = line.partition("|")[0]
            col = (line.partition("|")[2].partition("|")[0])
            arg1 = (line.partition("|")[2].partition("|")[2].partition("|")[0])
            arg2 = (line.partition("|")[2].partition("|")[2].partition("|")[2].partition("|")[0])
            arg3 = (line.partition("|")[2].partition("|")[2].partition("|")[2].partition("|")[2].partition("|")[0])
            output_file = output_file.replace(" ", "")
            
            col = col.replace(" ", "")
            arg1 = arg1.replace(" ", "")
            arg2 = arg2.replace(" ", "")
            arg3 = arg3.replace(" ", "")
            arg3 = arg3.replace("\n", "")

            
            
            print("---------------------------------------------------")
            print("Saving on file |" + output_file + "|")
            print("Searching on column |" + col + "| for |" + arg1 + "|" + arg2 + "|" + arg3 + "|")
            
            if len(output_files) != 0:
                flag = False
                for out in output_files:
                    if out == output_file:
                        print(str(out) + " == " + str(output_file))
                        out_df = pd.read_excel(output_file)
                        flag = True
                    else:
                        out_df = pd.DataFrame(columns=header)
                if not flag:
                    output_files.append(output_file)
                
            else:
                output_files.append(output_file)
                out_df = pd.DataFrame(columns=header)
            
            writer = pd.ExcelWriter(output_file, engine='xlsxwriter')

            column = 0

            for i in range(0, len(croped_df.columns)):
                c = croped_df.columns[i]
                if c.replace(" ", "") == col:
                    column = i

            if column == 0:
                print("No column named " + col)
                return
            
            for j in range(0, len(croped_df)):
                if arg1 in str(croped_df.iloc[j][croped_df.columns[column]]):
                    row = list()
                    for k in range(0, len(croped_df.columns)):
                        row.append(croped_df.iloc[j][croped_df.columns[k]])
                    out_df.loc[len(out_df.index)] = row
                    if j not in rows_to_delete:
                        rows_to_delete.append(j)
                if arg2 in str(croped_df.iloc[j][croped_df.columns[column]]):
                    row = list()
                    for k in range(0, len(croped_df.columns)):
                        row.append(croped_df.iloc[j][croped_df.columns[k]])
                    out_df.loc[len(out_df.index)] = row
                    if j not in rows_to_delete:
                        rows_to_delete.append(j)
                if arg3 in str(croped_df.iloc[j][croped_df.columns[column]]):
                    row = list()
                    for k in range(0, len(croped_df.columns)):
                        row.append(croped_df.iloc[j][croped_df.columns[k]])
                    out_df.loc[len(out_df.index)] = row
                    if j not in rows_to_delete:
                        rows_to_delete.append(j)
            
            out_df.to_excel(writer, index=False, header=True)
            writer.save()
            

    rows_to_delete = np.sort(rows_to_delete)[::-1]

    for i in range(0, len(rows_to_delete) - 1):
        print("deleting row " + str(rows_to_delete[i]))
        croped_df.drop(labels=rows_to_delete[i], axis=0, inplace=True)
    croped_df.to_excel(writer_croped, index=False, header=True)
    writer_croped.save()



if __name__ == '__main__':
    main()