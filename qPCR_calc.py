#!/usr/bin/env python3

'''
todo:
Get flow data
calculate 2^( (Flow_sample - Flow_gapdh) - (Sample_sample - Sample_gapdh) )
'''

import csv

DATA_FILE = 'sample_data.tsv'
CONTROL_TARGET = 'gapdh'
SAMPLE_NAME = 'Sample Name'
TARGET_NAME = 'Target Name'
CT = 'CT'
DELTA_CT = 'delta CT'
AVERAGE_CT = 'average CT'

class SampleClass:
    def __init__(self, CT):
        self.CTtotal = float(CT)
        self.count = 1.0
    
    def add(self, CT):
        self.CTtotal = self.CTtotal + float(CT)
        self.count = self.count + 1.0

    def getAverage(self):
        return self.CTtotal / self.count

    def getCount(self):
        return self.count



def main():
    data_dict = {} 
    with open(DATA_FILE) as tsvfile:
        reader = csv.DictReader(tsvfile, dialect='excel-tab') 
        for row in reader:
            sample = row[SAMPLE_NAME]
            target = row[TARGET_NAME]
            ct = row[CT]

            # check if target exists
            if target in data_dict:
                # check if sample exists in subdict
                if sample in data_dict[target]:
                    data_dict[target][sample].add(ct)
                else:
                    data_dict[target][sample] = SampleClass(ct)

            else:
                # initialize key with subdict of sample and SampleClass object
                data_dict[target] = {sample: SampleClass(ct)}

            print(row)
            print(row['Sample Name'])
            if data_dict[target][sample].getCount() == 3:
                print(data_dict[target][sample].getAverage())
    
    tsvfile.close()

    control_dict = data_dict[CONTROL_TARGET].copy()
    #print(control_dict)
    #print(data_dict)

    with open('out.tsv', 'w') as writetsv:
        outwriter = csv.writer(writetsv, dialect='excel-tab')
        outwriter.writerow([TARGET_NAME, SAMPLE_NAME, AVERAGE_CT, DELTA_CT])
        # loop through dictionary
        for target_key in data_dict:
            for sample_key in data_dict[target_key]:
                ct_average = data_dict[target_key][sample_key].getAverage()
                ct_delta = ct_average - control_dict[sample_key].getAverage()
                outwriter.writerow([target_key, sample_key, ct_average, ct_delta])
    #dialect='excel-tab')

main()