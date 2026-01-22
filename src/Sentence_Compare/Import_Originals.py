import os
import csv
import re
import logging
import datetime
import random

from pathlib import Path
from collections import Counter


class Import_Exception(Exception):
    """A custom exception class."""
    pass


class Import_Originals(object):
    '''Read and process data from CVS files in the given directory.'''
    
    # TODO:  Many methods in this class depend on a data structure in which
    #        the first element (record[0]) in an element of a list of records
    #        is the text to be processed.
    
    # TODO: Add copy method in assignment of records in get methods
    def __init__(self, path_name, truncate_at=0, Jabberwocky=False):
        '''Create object and save path to look in for CSV files.'''
        
        '''truncate_at and Jabberwocky are testing options. The default value of
        0 for truncate_at means do not truncate the records.
        '''
        
        self.path = Path(path_name)
        self.Jabberwocky = Jabberwocky
        if truncate_at > 0:
            self.num_recs = truncate_at
            self.rand_records = True
            if not self.rand_records:
                self.start_at = 100
                logging.debug(f'{self.num_recs} records selected starting at record {self.start_at}.\n')
        else:
            self.num_recs = 0
            
        # Adjust row entries for the worksheet and header offset.
        # 1 because the worksheet starts at 1 and not 0
        # 1 because of the header row in the converted sheet.
        self.excel_row_offset = 2

        

    def readin_csv_data(self, has_header=True, override=False):
        '''Create a list from csv files found in the given path.'''
        
        self.records = [] 
        self.converted_records = []
        self.sources = []
        self.sources_header = ['File_Name','Records_Read','Created','Modified','Accessed']
        try:
            for file_name in self.path.glob('*.csv'):
                self.read_cvs_file(file_name, has_header=has_header, override=override)
                
            # TODO: The following line depends on the layout of the CSV file.
            #       Might need changes when changing topics or sources.
                            
            # If debugging truncate to self.num_recs records.
            if __debug__:
                if self.num_recs > 0:
                    if self.Jabberwocky:
                        recs_to_keep = self.num_recs - 24
                    else:
                        recs_to_keep = self.num_recs
                    if len(self.converted_records) > recs_to_keep:
                        # Sets do not allow duplicates so if randint genereates 
                        # the same number twice, the second will be thrown away.
                        rand_recs = {}
                        while len(rand_recs) < recs_to_keep:
                            rec = random.randint(0, len(self.converted_records) - 1)
                            rand_recs[rec] = self.converted_records[rec]
                        self.converted_records = list(rand_recs.values())
                if self.Jabberwocky:
                    parent_dir = self.path.parent
                    # Add the desired file name
                    Jabberwocky_path = parent_dir / 'Jabberwocky.csv'
                    if Jabberwocky_path.exists():
                        self.read_cvs_file(Jabberwocky_path, has_header=has_header, override=override)                    
                        logging.debug(f'Sentence count after Jabberwocky: {len(self.converted_records)}')
            # Dedup converted records and sort both lists by the 'Volume' column.
            self.converted_records = self.dedupe_records(self.converted_records, 0)
            logging.debug(f'Converted sentence count after truncate: {len(self.converted_records)}')
            self.records = sorted(self.records, key=lambda x: x[4], reverse=True)
            self.converted_records = sorted(self.converted_records, key=lambda x: x[2], reverse=True)
            # Add a row to track original order of records
            for i, row in enumerate(self.converted_records):
                row.extend([i])
            
        except Import_Exception as e:
            print(f'Import Exception:\n{e}')
        except Exception as e:
            print(f"Unexpected error:\n{e}")
        if not self.records:
            raise Import_Exception('No CSV files found or records read.')
        
    def read_cvs_file(self, file_path, has_header=True, override=False):
        '''Read in the rows from the specified CSV file and process them.'''
        
        '''Two lists are created the records list and the converted records list.
        If has_header is True the first files header will be removed from the
        records, saved and compared to subsiquent file headers. A mismatch will
        result in an exception unless override=True is specified.
        '''
        
        '''The converted records sentences have been stripped of any leading 
        and trailing spaces and the numeric fields have been converted from 
        strings to numbers. After this any duplicate records are removed. The 
        internals of the sentences are not processed and may contain punctuation 
        etc.
        '''
        
        # TODO: The following lines depends on the layout of the CSV file.
        #       Might need changes when changing topics or sources.        
        try:
            # Open the CSV file
            with open(file_path, 'r') as file:
                # Create a CSV reader object
                reader = csv.reader(file)
                # Iterate over each row in the CSV file
                rows = []
                converted_rows = []
                for row in reader:
                    rows.append(row)
                    # Just take the sentence and the difficulty and volume fields.
                    conv_row = [row[1], row[3], row[4]]
                    # TODO: Find a better way to skip the column names row.
                    if converted_rows:
                        # Remove leading and trailing spaces from the sentence 
                        # and leave everything 'inside' the sentence the same.
                        conv_row[0] = conv_row[0].strip()
                        # Convert Difficulty and Volume to ints and trap for empty strings
                        if conv_row[1].isnumeric():
                            conv_row[1] = int(conv_row[1])
                        else:
                            conv_row[1] = 0
                        if conv_row[2].isnumeric():    
                            conv_row[2] = int(conv_row[2])
                        else:
                            conv_row[2] = 0
                    converted_rows.append(conv_row)
                if has_header:
                    # Remove headder
                    if not hasattr(self, 'original_header'):
                        self.original_header = rows.pop(0)
                        self.converted_header = converted_rows.pop(0)
                        self.converted_header.extend(['Sort_Order'])
                    elif not override:
                        new_header = rows.pop(0)
                        converted_rows.pop(0)
                        for ndx, value in enumerate(self.original_header):
                            if not new_header[ndx] == value:
                                raise Import_Exception('The Headers do not match. You can pass override=True to readin_cvs_data() to ignore this.')
                    else:
                        rows.pop(0)
                        converted_rows.pop(0)
                # Save this information for the Sources sheet of the workbook.
                file_stats = os.stat(file_path)
                # Convert the timestamps to datetime objects
                created = datetime.datetime.fromtimestamp(file_stats.st_ctime)
                modified = datetime.datetime.fromtimestamp(file_stats.st_mtime)
                accessed = datetime.datetime.fromtimestamp(file_stats.st_atime)
                created_str = created.strftime("%Y-%m-%d %H:%M:%S")
                modified_str = modified.strftime("%Y-%m-%d %H:%M:%S")
                accessed_str = accessed.strftime("%Y-%m-%d %H:%M:%S")
                
                file_name = os.path.basename(file_path)
                num_recs = len(rows)
                self.sources.append([file_name, num_recs, created_str, modified_str, accessed_str])

                self.records.extend(rows)
                self.converted_records.extend(converted_rows)

        except IOError as e:
            print(f'Error opening the file: {file_name}:\n{e}')
        except csv.Error as e:
            print(f'Error reading the CSV file: {file_name}:\n{e}')
        except Exception as e:
            print(f"Unexpected error processing {file_name}:\n{e}")


    def get_original_header(self):
        ''' Return the header of the original file.''' 
        
        if hasattr(self, 'original_header'):
            return self.original_header
        else:
            logging.info('No header information available.\n')
            return None
    

    def get_converted_header(self):
        '''Only the header values for the converted rows are returned.'''
        
        if hasattr(self, 'converted_header'):
            return self.converted_header
        else:
            logging.info('No converted_header information available.\n')
            return None
    

    def print_records(self):
        '''Print the records to the display line by line.'''
        
        ''' May take a while depending on the number of records!'''
        
        if hasattr(self, 'records'):
            print(f'\nRecords:')
            for rec in self.records:
                print(rec)
        else:
            logging.info('No records have been created.\n')

    def get_records(self, header=True):
        '''Returns the raw records, optionally with the header information.'''
        
        if hasattr(self, 'records'):
            recs = self.records.copy()
            if header:
                recs.insert(0, self.get_original_header())
            return recs
        else:
            logging.info('No records have been created.\n')
            return None       

    def get_sentences(self):
        '''Return just the sentence information from the raw records.'''
        
        # TODO: The following lines depends on the layout of the CSV file.
        if hasattr(self, 'records'):
            sentences = [rec[1] for rec in self.records]
            return sentences
        else:
            logging.info('No records have been created.\n')
            return None      
        

    def print_deduped_records(self):
        ''' Remove any duplicates in the raw sentences and prints the results.'''
        
        if hasattr(self, 'records'):
            print(f'\nSentence records deduped:')
            for rec in self.dedupe_records(self.records, 1):
                print(rec)
        else:
            logging.info('No records have been created.\n')

    def get_deduped_records(self, header=True):
        ''' Remove duplicates and return the resultant list of records,'''
        
        if hasattr(self, 'records'):
            recs = self.dedupe_records(self.records, 1)
            if header:
                recs.insert(0, self.get_original_header())
            return recs
        else:
            logging.info('No records have been created.\n')
            return None       


    def get_deduped_sentences(self):
        ''' Remove duplicate sentences in the records and return the result.'''
        
        if hasattr(self, 'records'):
            sentences = [rec[1] for rec in self.dedupe_records(self.records, 1)]
            return sentences
        else:
            logging.info('No records have been created.\n')
            return None      
        
    def print_clean_records(self):
        '''Print the records with their sentences cleaned.'''
        
        '''The sentence fields are purged of extra spaces and punctuation, then
        any duplicates are removed and the results are printed out.
        '''
        
        if hasattr(self, 'records'):
            print(f'\nCleaned records:')
            for rec in self.clean_records(self.records):
                print(rec)
        else:
            logging.info('No records have been created.\n')

    def get_clean_records(self, header=True):
        '''Return the records with their sentences cleaned.'''
    
        '''The raw sentence fields are purged of extra spaces and punctuation,
        then any duplicates are removed and the resultant records are returned.
        '''
        
        if hasattr(self, 'records'):
            recs = self.clean_records(self.records)
            if header:
                recs.insert(0, self.get_original_header())
            return recs
        else:
            logging.info('No records have been created.\n')
            return None       

    def get_clean_sentences(self):
        '''Returns the sentences from the records after they are cleaned.'''
        
        '''The raw sentence fields are purged of extra spaces and punctuation,
        then duplicates are removed and the resultant sentences are returned.
        '''
        
        if hasattr(self, 'records'):
            sentences = [rec[1] for rec in self.clean_records(self.records)]
            return sentences
        else:
            logging.info('No records have been created.\n')
            return None      

            
    def print_converted_records(self):
        '''Print the list of converted records to the display.'''

        if hasattr(self, 'converted_records'):
            print(f'\nConverted records:')
            for rec in self.converted_records:
                print(rec)
        else:
            logging.info('No records have been created.\n')

    def get_converted_records(self, header=True):
        '''Return the list of converted records.'''
        
        if hasattr(self, 'converted_records'):
            recs = self.converted_records.copy()
            if header:
                recs.insert(0, self.get_converted_header())
            return recs
        else:
            logging.info('No records have been created.\n')
            return None       

    def get_converted_records_sheet(self):
        '''Return the list of converted records.'''
        
        if hasattr(self, 'converted_records'):
            recs = self.converted_records.copy()
            recs = [[r[0], r[1], r[2], r[3] + self.excel_row_offset] for r in recs]
            recs.insert(0, self.get_converted_header())
            return recs
        else:
            logging.info('No records have been created.\n')
            return None       

    def get_converted_sentences(self):
        '''Return the list of sentences from the converted records.'''
        
        # TODO: The following lines depends on the layout of the converted 
        # records that were created in the read_cvs_file() method.
        if hasattr(self, 'converted_records'):
            sentences = [rec[0] for rec in self.converted_records]
            return sentences
        else:
            logging.info('No records have been created.\n')
            return None              


    def print_clean_converted_records(self):
        '''Clean the sentences in the converted records and print them.'''
        
        '''Cleaning will remove excess spaces, punctuation, and non-alphanumeric
        characters.
        '''

        if hasattr(self, 'converted_records'):
            print(f'\nCleaned converted records:')
            for rec in self.clean_records(self.converted_records):
                print(rec)
        else:
            logging.info('No records have been created.\n')

    def get_clean_converted_records(self, header=True):
        '''Clean the sentences in the records and return the cleaned records.'''

        if hasattr(self, 'converted_records'):
            recs = self.clean_records(self.converted_records)
            if header:
                recs.insert(0, self.get_converted_header())
            return recs
        else:
            logging.info('No records have been created.\n')
            return None       

    def get_cleaned_converted_sentences(self):
        '''Clean the converted record's sentences and return the sentences.'''
        
        if hasattr(self, 'converted_records'):
            sentences = [rec[0] for rec in self.clean_records(self.converted_records)]
            return sentences
        else:
            logging.info('No records have been created.\n')
            return None      
        

    def print_sources(self):
        '''Print out information about the files that were read in.'''
        
        if hasattr(self, 'sources'):
            print(f'\nSources:\n{self.sources_header}')
            for source in self.sources:
                print(f'{source}')
        else:
            logging.info('No files have been read.\n')

    
    def get_sources(self, header=True):
        '''Return a list of information about files that were read.'''

        rtn_mtrx = []
        if hasattr(self, 'sources'):
            if header:
                rtn_mtrx.append(self.sources_header)
            rtn_mtrx.extend(self.sources)
            return rtn_mtrx
        else:
            logging.info('No files have been read.\n')
            return None

    def dedupe_records(self, records, sentence_field):
        ''' Go through the list of records and remove any duplicate records.'''
        
        '''A cleaned version of each sentence is compared to set of unique 
        cleaned sentences and only one is kept.
        '''
        
        if records:
            # Deduplicate the cleaned sentences
            logging.debug(f'Sentence count before dedupe: {len(records)}')
            
            # This give a list of clean deduped sentences.
            clean_sentences = list(set([re.sub(r'[^\w\s]', '', record[sentence_field].strip().casefold()) for record in records]))
            
            logging.debug(f'Sentence count after dedupe: {len(clean_sentences)}\n')

            skip_list = [False] * len(clean_sentences)
            # Create a set to store deduped clean sentences along with their associated booleans
            sentence_set = set(zip(clean_sentences, skip_list))

            # List to store raw sentences that match deduped cleaned sentences and need to be updated
            matched_records = []

            # Iterate over the raw sentences
            for ndx, record in enumerate(records):
                key = re.sub(r'[^\w\s]', '', record[sentence_field].strip().casefold())
                for i, (cleaned_sentence, skip) in enumerate(sentence_set):
                    if key == cleaned_sentence and not skip:
                        # If match found and boolean is False, add the raw sentence to the list
                        matched_records.append(record)
                        # Update the boolean to True
                        sentence_set.remove((cleaned_sentence, skip))
                        sentence_set.add((cleaned_sentence, True))
                        # Break out of the inner loop after finding the first match.
                        # Any subsequent matches will be thrown away
                        break
                    
            if __debug__:
                dedupe_len = len(sentence_set)
                matched_len = len(matched_records)
                assert dedupe_len == matched_len, f'Dedup failed list lengths do not match {dedupe_len} != {matched_len}'

            return matched_records
        else:
            raise Import_Exception('Empty records list passed to dedupe_records.')
 
 
    def clean_records(self, records):
        '''Clean the given records sentences and return a list of the records.'''
        
        '''For the sentence in each record: remove leading and trailling 
        spaces, shift to lower case, and remove extra spaces and 
        non-alphanumeric characters.'''
        
        if records:
            if records is self.records:
                sentence_field = 1
            else:
                sentence_field = 0
            if sentence_field == 0:
                records_clean = [[re.sub(r'[^\w\s]', '', row[0].strip().casefold())] for row in records]
                for i, row in enumerate(records_clean):
                    row.extend(records[i][1:])
            else:
                records_clean = [[row[0], re.sub(r'[^\w\s]', '', row[1].strip().casefold())] for row in records]
                for i, row in enumerate(records_clean):
                    row.extend(records[i][2:])
               
            records_clean = self.dedupe_records(records_clean, sentence_field)
         
            return records_clean
        else:
            raise Import_Exception('Empty records list passed to clean_records.')

            















    
