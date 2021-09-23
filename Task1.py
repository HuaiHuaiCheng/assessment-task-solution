import glob  # glob (short for global) is used to return all file paths that match a specific pattern.
import zipfile
import pandas as pd
from natsort import natsorted
filename = 'combine files.zip'  # name of the zip file
zip_file = glob.glob(filename)  # search the file with name 'filename'
zf = zipfile.ZipFile(zip_file[0])   # read zip file
csv_files = zf.infolist()

dfs = []
aa = natsorted(zf.namelist())  # sort the namelist
for f in aa:
    ff = pd.read_csv(zf.open(f),header=0)  # read each csv file
    dfs.append(ff)  # append each csv file into a list one after another
    df = pd.concat(dfs,ignore_index=True)  # combine all csv files together
    df.drop_duplicates()  # drop the duplicate header
    df.to_csv('Task1.csv',index=False)  # save the 10 files data into a new csv file and ignore index
