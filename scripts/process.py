import csv
import urllib
import datautil.tabular.xls as xlstab

date = "oct2015"
source = 'http://www.genome.gov/pages/der/sequencing_costs_'+date+'.xlsx?force_download=true'
in_path = 'archive/sequencing_costs_'+date+'.xlsx'
out_path = 'data/sequencing_costs.csv'
def execute():
    urllib.urlretrieve(source, in_path)
    records = []

    reader = xlstab.XlsReader()
    tabdata = reader.read(open(in_path), sheet_index=0)
    records = tabdata.data
    header = ['Date', 'Cost per Mb', 'Cost per Genome']
    records = records[1:]
    for x in records:
        x[0] = str(x[0].year)+'-'+str(x[0].month)
    print (records)

    writer = csv.writer(open(out_path, 'w'), lineterminator='\n')
    writer.writerow(header)
    writer.writerows(records)

execute()
