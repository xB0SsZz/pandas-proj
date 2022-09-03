import os
import pandas as pd


def main():
    for filename in os.listdir(os.getcwd()):
        if filename.endswith('.xlsx'):
            df = pd.read_excel(filename)
            print(filename + ": " + str(len(df.index)) + " records.")


if __name__ == '__main__':
    main()