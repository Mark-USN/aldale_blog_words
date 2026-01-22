import os
import time
import logging
from Import_Originals import Import_Originals
from Group_Transform import Group_Transform
from Excel_Writer import Excel_Writer

if __debug__:
    # Configure logging to display info or debug messages
    logging.basicConfig(level=logging.INFO)
    # logging.basicConfig(level=logging.DEBUG)
    # logging.debug('Setting the logging level to DEBUG.')


def main():
    ''' Actually put the pieces together to create the workbooks.'''
    
    if __debug__:
        prog_start = time.perf_counter()
    print('Starting the program.')
        
    if os.path.expanduser("~").__contains__('mhenw'):
        originals_path = 'C:\\Users\\mhenw\\source\\repos\\Sentence_Compare\\data_folder_original'
        workbook_path = 'C:\\Users\\mhenw\\source\\repos\\Sentence_Compare\\data_folder_result\\Sentence_Compare.xlsx'
    else:
        originals_path = 'C:\\Users\\aldal\\Desktop\\Words_Projects\\Sentence_Compare\\data_folder_original'
        workbook_path = 'C:\\Users\\aldal\\Desktop\\Words_Projects\\Sentence_Compare\\data_folder_result\\Sentence-Compare.xlsx'
       
    if __debug__:
        start_time = time.perf_counter()
        
    ''' Read in the CSV files and convert the fields from text as needed.'''
    
    # DANGER WILL ROBERTSON!: For testing purposes you can truncate the number 
    # of records to process and add the Jabberworcy poem to inject nonsence 
    # text. The Jabberworcy.csv must be in the parent directory of the path 
    # given to Import_Originals.     

    truncate_at=350


    # data_importer = Import_Originals(originals_path, truncate_at=truncate_at, Jabberwocky=True)
    data_importer = Import_Originals(originals_path)
    
    data_importer.readin_csv_data()
    
    # Use one of these for Group_Transform initialization.
    converted_records = data_importer.get_converted_records(header=False) 
    cleaned_converted_records = data_importer.get_clean_converted_records(header=False)

    # data_importer.print_sources()

    print('Completed reading in CSV files.\n')
    
    if __debug__:
        end_time = time.perf_counter()
        elapsed_time = end_time - start_time
        logging.debug(f'Reading in initial information took {elapsed_time:.4f} seconds.\n')

    ''' Use records obtained above to compute the similarity between sentences. 
    find the best matches and use these columns as topics.  Finally any columns
    containing consolidate_level or fewer rows are moved to another column.
    '''

    # DANGER WILL ROBERTSON!: These values can be changed after the 
    # Group_Transform object is created but many of the methods must be rerun
    # in order to see the effects of these changes.
    # The threshold value does not have alot of effect as it can be changed in
    # the workbook's Summary sheet.
        
    threshold=0.75
    consolidate_level = 8

    if __debug__:
        creation_time = time.perf_counter()
    grp_trans = Group_Transform(converted_records, consolidate_level=consolidate_level, threshold=threshold)
    if __debug__:
        end_time = time.perf_counter()
        elapsed_time = end_time - creation_time
        logging.debug(f'Group_Transformation creation took {elapsed_time:.4f} seconds.')


    if __debug__:
        start_time = time.perf_counter()
    grp_trans.compute_embeddings()
    if __debug__:
        end_time = time.perf_counter()
        elapsed_time = end_time - start_time
        logging.debug(f'Group_Transformation compute embeddings took {elapsed_time:.4f} seconds.')

    if __debug__:
        start_time = time.perf_counter()
    grp_trans.compute_cosines() 
    if __debug__:
        end_time = time.perf_counter()
        elapsed_time = end_time - start_time
        logging.debug(f'Group_Transformation compute cosines took {elapsed_time:.4f} seconds.')

    if __debug__:
        start_time = time.perf_counter()
    grp_trans.create_lower_triangle_matrix()
    if __debug__:
        end_time = time.perf_counter()
        elapsed_time = end_time - start_time
        logging.debug(f'Group_Transformation create lower triangle took {elapsed_time:.4f} seconds.')

    if __debug__:
        start_time = time.perf_counter()
    grp_trans.sort_lower_triangle_matrix()
    if __debug__:
        end_time = time.perf_counter()
        elapsed_time = end_time - start_time
        logging.debug(f'Group_Transformation sort lower triangle took {elapsed_time:.4f} seconds.')

    if __debug__:
        start_time = time.perf_counter()
    grp_trans.create_best_matrix()
    if __debug__:
        end_time = time.perf_counter()
        elapsed_time = end_time - start_time
        logging.debug(f'Group_Transformation create best matrix took {elapsed_time:.4f} seconds.')

    if __debug__:
        start_time = time.perf_counter()
    grp_trans.create_consolidated_matrix()
    if __debug__:
        end_time = time.perf_counter()
        elapsed_time = end_time - start_time
        logging.debug(f'Group_Transformation create consolidated matrix took {elapsed_time:.4f} seconds.')

    if __debug__:
        start_time = time.perf_counter()
    grp_trans.create_topic_matrix()
    if __debug__:
        end_time = time.perf_counter()
        elapsed_time = end_time - start_time
        logging.debug(f'Group_Transformation create topic matrix took {elapsed_time:.4f} seconds.')

    print('Completed processing the information passed in from Import_Originals().\n')
    
    if __debug__:
        end_time = time.perf_counter()
        elapsed_time = end_time - creation_time
        logging.debug(f'Processing initial information took {elapsed_time:.4f} seconds.\n')

    ''' The following block - actually now only one method - does not need to 
    be run in production (release) mode.
    '''

    if __debug__:
        start_time = time.perf_counter()
        logging.debug('Starting test methods.\n')
        grp_trans.check_all_lookup_tables()
        logging.debug('Finished Lookup Table check.')
        elapsed_time = end_time - start_time
        
        print('Completed Group_Transformer integrity tests.\n')
        
        logging.debug(f'Test methods took {elapsed_time:.4f} seconds.\n')

    logging.debug('Starting to create and save executive summary spreadsheets.\n')
    
    ''' Ensure the file name being used is unique. '''
    
    # Split the file path into directory, base filename, and extension
    _, filename = os.path.split(workbook_path)
    original_base_name, _ = os.path.splitext(filename)


    base_name_modifier = 0
    while os.path.exists(workbook_path):
       
        base_name_modifier -= 1

        # Split the file path into directory, base filename, and extension
        dir_path, old_filename = os.path.split(workbook_path)
        base_name, extension = os.path.splitext(old_filename)
    
        # Construct the new file path with the new filename
        new_filename = original_base_name + f'{base_name_modifier}'
        workbook_path = os.path.join(dir_path, new_filename + extension)
    
    # Split the file path into directory, base filename, and extension
    dir_path, old_filename = os.path.split(workbook_path)
    base_name, extension = os.path.splitext(old_filename)
    if base_name_modifier < 0:
        # Construct the new file path with the new filename
        new_filename = original_base_name + '_Executive_Summary' +  f'{base_name_modifier}'
    else:
        new_filename = original_base_name + '_Executive_Summary'
        
    exe_sum_file_path = os.path.join(dir_path, new_filename + extension)

    ''' Create and save Executive Summary Sheet. '''
    
    if __debug__:
        writer_start_time = time.perf_counter()
    writer = Excel_Writer(exe_sum_file_path)
    
    # Sheets for release version
    if __debug__:
        start_time = time.perf_counter()
    org_data = data_importer.get_records(header=True)
    org_header = org_data.pop(0)
    writer.create_original_sheet(org_data, org_header) 
    if __debug__:
        end_time = time.perf_counter()
        elapsed_time = end_time - start_time
        logging.debug(f'Creating Original Sheet took {elapsed_time:.4f} seconds.')

    if __debug__:
        start_time = time.perf_counter()
    cnvtd_data = data_importer.get_converted_records_sheet()
    cnvtd_header = cnvtd_data.pop(0)
    writer.add_sheet(cnvtd_data, 'Converted', header=cnvtd_header, slant_header=False)
    if __debug__:
        end_time = time.perf_counter()
        elapsed_time = end_time - start_time
        logging.debug(f'Creating Converted Sheet took {elapsed_time:.4f} seconds.')

    if __debug__:
        start_time = time.perf_counter()
    aldos_fltrd_data, aldos_header = grp_trans.get_aldos_sheet()
    writer.add_aldos_sheet(aldos_fltrd_data, 'Aldos', header=aldos_header, slant_header=False)
    if __debug__:
        end_time = time.perf_counter()
        elapsed_time = end_time - start_time
        logging.debug(f'Creating Aldos Sheet took {elapsed_time:.4f} seconds.')

    if __debug__:
        start_time = time.perf_counter()
    topic_data, topic_header = grp_trans.get_topic_sheet()
    writer.add_sheet(topic_data, 'Topics', cell_width=50, header=topic_header, slant_header=False)
    if __debug__:
        end_time = time.perf_counter()
        elapsed_time = end_time - start_time
        logging.debug(f'Creating Topics Sheet took {elapsed_time:.4f} seconds.')

    if __debug__:
        start_time = time.perf_counter()
    sentence_data, sentence_header = grp_trans.get_consolidated_sentences_sheet()
    writer.add_sheet(sentence_data, 'Sentences', header=sentence_header, cell_width=50, slant_header=False)
    if __debug__:
        end_time = time.perf_counter()
        elapsed_time = end_time - start_time
        logging.debug(f'Creating Sentences Sheet took {elapsed_time:.4f} seconds.')

    if __debug__:
        start_time = time.perf_counter()
    cnsldt_data, cnsldt_header = grp_trans.get_consolidated_cosine_sheet()
    writer.add_sheet(cnsldt_data, 'Consolidated', header=cnsldt_header)
    if __debug__:
        end_time = time.perf_counter()
        elapsed_time = end_time - start_time
        logging.debug(f'Creating Consolidated Sheet took {elapsed_time:.4f} seconds.')

    if __debug__:
        start_time = time.perf_counter()
    cnsldt_trck_data, cnsldt_trck_header = grp_trans.get_consolidated_tracking_sheet()
    writer.add_sheet(cnsldt_trck_data, 'Consolidated_Tracking', header=cnsldt_trck_header)
    if __debug__:
        end_time = time.perf_counter()
        elapsed_time = end_time - start_time
        logging.debug(f'Creating Consolidated_Tracking Sheet took {elapsed_time:.4f} seconds.')
        
    if __debug__:
        start_time = time.perf_counter()
    source_data = data_importer.get_sources(header=True)
    if source_data is not None:
        source_header = source_data.pop(0)
        num_ini_topics = grp_trans.get_initial_topic_count()
        num_con_topics = grp_trans.get_consolidated_topic_count()
        con_fact = 1 - float(num_con_topics)/float(num_ini_topics)
        source_data.append(['', 'Threshold used: ', grp_trans.get_threshold(), '', '' ])
        source_data.append(['', 'Consolidation level: ', grp_trans.get_consolidation_level(), '', '' ])
        source_data.append(['', 'Total Records: ', grp_trans.get_records_count(), '', '' ])
        source_data.append(['', 'Initial Topics: ', num_ini_topics, '', '' ])
        source_data.append(['', 'Consolidated Topics: ', num_con_topics, '', '' ])
        source_data.append(['', 'Consolidation factor: ', con_fact, '', '' ])

        writer.add_sheet(source_data, 'Summary', header=source_header, slant_header=False)
    if __debug__:
        end_time = time.perf_counter()
        elapsed_time = end_time - start_time
        logging.debug(f'Creating Summary Sheet took {elapsed_time:.4f} seconds.')
        
    if __debug__:
        start_time = time.perf_counter()
    writer.write_workbook()
    if __debug__:
        end_time = time.perf_counter()
        elapsed_time = end_time - start_time
        logging.debug(f'Writing out the workbook took {elapsed_time:.4f} seconds.\n')
        writer_time = end_time - writer_start_time
        logging.debug(f'Total time for excel operation was {writer_time:.4f} seconds.\n')
        
    print('Completed creating and writing out the executive summary workbook.\n')
    
    ''' Create and save all Sheets. '''
    
    print('Starting to create and save the complete workbook.\n')
    
    if __debug__:
        writer_start_time = time.perf_counter()
    writer = Excel_Writer(workbook_path)
    # Sheets for release version
    if __debug__:
        start_time = time.perf_counter()
    org_data = data_importer.get_records(header=True)
    org_header = org_data.pop(0)
    writer.create_original_sheet(org_data, org_header) 
    if __debug__:
        end_time = time.perf_counter()
        elapsed_time = end_time - start_time
        logging.debug(f'Creating Original Sheet took {elapsed_time:.4f} seconds.')

    if __debug__:
        start_time = time.perf_counter()
    cnvtd_data = data_importer.get_converted_records_sheet()
    cnvtd_header = cnvtd_data.pop(0)
    writer.add_sheet(cnvtd_data, 'Converted', header=cnvtd_header, slant_header=False)
    if __debug__:
        end_time = time.perf_counter()
        elapsed_time = end_time - start_time
        logging.debug(f'Creating Converted Sheet took {elapsed_time:.4f} seconds.')

    if __debug__:
        start_time = time.perf_counter()
    aldos_fltrd_data, aldos_header = grp_trans.get_aldos_sheet()
    writer.add_aldos_sheet(aldos_fltrd_data, 'Aldos', header=aldos_header, slant_header=False)
    if __debug__:
        end_time = time.perf_counter()
        elapsed_time = end_time - start_time
        logging.debug(f'Creating Aldos Sheet took {elapsed_time:.4f} seconds.')

    if __debug__:
        start_time = time.perf_counter()
    topic_data, topic_header = grp_trans.get_topic_sheet()
    writer.add_sheet(topic_data, 'Topics', cell_width=50, header=topic_header, slant_header=False)
    if __debug__:
        end_time = time.perf_counter()
        elapsed_time = end_time - start_time
        logging.debug(f'Creating Topics Sheet took {elapsed_time:.4f} seconds.')

    if __debug__:
        start_time = time.perf_counter()
    sentence_data, sentence_header = grp_trans.get_consolidated_sentences_sheet()
    writer.add_sheet(sentence_data, 'Sentences', header=sentence_header, cell_width=50, slant_header=False)
    if __debug__:
        end_time = time.perf_counter()
        elapsed_time = end_time - start_time
        logging.debug(f'Creating Sentences Sheet took {elapsed_time:.4f} seconds.')

    if __debug__:
        start_time = time.perf_counter()
    cos_data, cos_header = grp_trans.get_cosine_sheet()
    writer.add_sheet(cos_data, 'Cosines', header=cos_header, titles=True)
    if __debug__:
        end_time = time.perf_counter()
        elapsed_time = end_time - start_time
        logging.debug(f'Creating Cosines Sheet took {elapsed_time:.4f} seconds.')

    if __debug__:
        start_time = time.perf_counter()
    tri_data, tri_header = grp_trans.get_lower_triangle_sheet()
    writer.add_sheet(tri_data, 'Triangle', header=tri_header, titles=True)
    if __debug__:
        end_time = time.perf_counter()
        elapsed_time = end_time - start_time
        logging.debug(f'Creating Triangle Sheet took {elapsed_time:.4f} seconds.')

    if __debug__:
        start_time = time.perf_counter()
    sorted_data, srtd_header = grp_trans.get_sorted_cosine_sheet()
    writer.add_sheet(sorted_data, 'Sorted', header=srtd_header)
    if __debug__:
        end_time = time.perf_counter()
        elapsed_time = end_time - start_time
        logging.debug(f'Creating Sorted Sheet took {elapsed_time:.4f} seconds.')

    if __debug__:
        start_time = time.perf_counter()
    look_up_data, lu_header = grp_trans.get_sorted_lookup_sheet()
    writer.add_sheet(look_up_data, 'Sorted_Lookup', header=lu_header)
    if __debug__:
        end_time = time.perf_counter()
        elapsed_time = end_time - start_time
        logging.debug(f'Creating Sorted_Lookup Sheet took {elapsed_time:.4f} seconds.')

    if __debug__:
        start_time = time.perf_counter()
    bst_data, bst_header = grp_trans.get_best_cosine_sheet()
    writer.add_sheet(bst_data, 'Best', header=bst_header)
    if __debug__:
        end_time = time.perf_counter()
        elapsed_time = end_time - start_time
        logging.debug(f'Creating Best Sheet took {elapsed_time:.4f} seconds.')

    if __debug__:
        start_time = time.perf_counter()
    best_trck, bst_trck_hdr = grp_trans.get_best_tracking_sheet()
    writer.add_sheet(best_trck, 'Best_Tracking', header=bst_trck_hdr)
    if __debug__:
        end_time = time.perf_counter()
        elapsed_time = end_time - start_time
        logging.debug(f'Creating Best_Tracking Sheet took {elapsed_time:.4f} seconds.')

    if __debug__:
        start_time = time.perf_counter()
    cnsldt_data, cnsldt_header = grp_trans.get_consolidated_cosine_sheet()
    writer.add_sheet(cnsldt_data, 'Consolidated', header=cnsldt_header)
    if __debug__:
        end_time = time.perf_counter()
        elapsed_time = end_time - start_time
        logging.debug(f'Creating Consolidated Sheet took {elapsed_time:.4f} seconds.')

    if __debug__:
        start_time = time.perf_counter()
    cnsldt_trck_data, cnsldt_trck_header = grp_trans.get_consolidated_tracking_sheet()
    writer.add_sheet(cnsldt_trck_data, 'Consolidated_Tracking', header=cnsldt_trck_header)
    if __debug__:
        end_time = time.perf_counter()
        elapsed_time = end_time - start_time
        logging.debug(f'Creating Consolidated_Tracking Sheet took {elapsed_time:.4f} seconds.')

    if __debug__:
        start_time = time.perf_counter()
    source_data = data_importer.get_sources(header=True)
    if source_data is not None:
        source_header = source_data.pop(0)
        num_ini_topics = grp_trans.get_initial_topic_count()
        num_con_topics = grp_trans.get_consolidated_topic_count()
        con_fact = 1 - float(num_con_topics)/float(num_ini_topics)
        source_data.append(['', 'Threshold used: ', grp_trans.get_threshold(), '', '' ])
        source_data.append(['', 'Consolidation level: ', grp_trans.get_consolidation_level(), '', '' ])
        source_data.append(['', 'Total Records: ', grp_trans.get_records_count(), '', '' ])
        source_data.append(['', 'Initial Topics: ', num_ini_topics, '', '' ])
        source_data.append(['', 'Consolidated Topics: ', num_con_topics, '', '' ])
        source_data.append(['', 'Consolidation factor: ', con_fact, '', '' ])

        writer.add_sheet(source_data, 'Summary', header=source_header, slant_header=False)
    if __debug__:
        end_time = time.perf_counter()
        elapsed_time = end_time - start_time
        logging.debug(f'Creating Summary Sheet took {elapsed_time:.4f} seconds.')
                
       
    if __debug__:
        start_time = time.perf_counter()
    writer.write_workbook()
    if __debug__:
        end_time = time.perf_counter()
        elapsed_time = end_time - start_time
        logging.debug(f'Writing out the workbook took {elapsed_time:.4f} seconds.\n')
        writer_time = end_time - writer_start_time
        logging.debug(f'Total time for excel operation was {writer_time:.4f} seconds.\n')

    print('Completed creating and writing out the complete workbook.\n')

    # grp_trans.set_threshold(original_threshold)
    
    if __debug__:
        end_time = time.perf_counter()
        elapsed_time = end_time - prog_start
        logging.debug(f'The program took {elapsed_time:.4f} seconds.')
    
    
if __name__ == "__main__":
    main()
