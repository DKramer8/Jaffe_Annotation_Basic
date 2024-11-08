# ------------------------------------------------------------------------------------
# ------------------------------------------------------------------------------------
# Imports
# ------------------------------------------------------------------------------------
# ------------------------------------------------------------------------------------
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.patches as patches
import os
import pagexml.helper.text_helper as text_helper
import io
from pagexml.helper.file_helper import  read_page_archive_files
from pagexml import parser as pxml_parser
from zipfile import ZipFile

# ------------------------------------------------------------------------------------
# ------------------------------------------------------------------------------------
# Globals
# ------------------------------------------------------------------------------------
# ------------------------------------------------------------------------------------
DIR_PATH = os.getcwd()
INPUT_PATH = DIR_PATH + f'\\input'
OUTPUT_PATH = DIR_PATH + f'\\output'

TEXTREGION_THRESHOLD_MIN_WIDTH = 1000
TEXTREGION_THRESHOLD_MIN_HEIGHT = 1000
HEADERREGION_THRESHOLD_MIN_X_COORD = 500
HEADERREGION_THRESHOLD_MAX_X_COORD = 2000
HEADERREGION_THRESHOLD_MAX_H_COORD = 500
NUMBERREGION_THRESHOLD_MAX_X_COORD = 750
NUMBERREGION_THRESHOLD_MAX_W_COORD = 500
NUMBERREGION_THRESHOLD_MIN_X_COORD = 2250

REGEST_THRESHOLD_MIN_X_COORD_LEFT = 750
PLACE_THRESHOLD_MIN_X_COORD_LEFT =  500
REGEST_THRESHOLD_MIN_X_COORD_RIGHT = 600
PLACE_THRESHOLD_MIN_X_COORD_RIGHT =  350

NEW_POPE_THRESHOLD_Y_COOD = 300

# ------------------------------------------------------------------------------------
# ------------------------------------------------------------------------------------
# Functions
# ------------------------------------------------------------------------------------
# ------------------------------------------------------------------------------------
def console_welcome():
    '''
        Prints out the welcome text and asks for input format

        :return: The entered input in lower format
    '''

    print('----------------------------------------------------------------------------------------------------------------------')
    print('Welcome to the classification script for Jaffé OCR')
    print('----------------------------------------------------------------------------------------------------------------------')
    print('Do you want to classify a specific PageXML file or a whole Zip Archive of PageXML? (single/zip)')
    i = input()
    print('----------------------------------------------------------------------------------------------------------------------')

    return i.lower()

def new_dictionaries():
    '''
        Helper function to easily create new empty dictionaries of the necessary format.

        :return: textRegion_dict, unknown_region_dict, final_regest_dict
    '''

    textRegion_dict = {
        'x': [],
        'y': [],
        'w': [],
        'h': [],
        'type': [],
        'text': []
    }
    unknown_region_dict = {
        'x': [],
        'y': [],
        'h': [],
        'w': []
    }
    final_regest_dict = {
        'pope': [],
        'date': [],
        'place': [],
        'number': [],
        'text': [],
        'page_nr': []
    }

    return textRegion_dict, unknown_region_dict, final_regest_dict

def add_coords_to_dict(dict, coords):
    '''
        Adds the given coords to the respective list of the given dictionary. Coords and dict does have specific conditions.

        :param dict: The dictionary to which the coords get added to. A dictionary with a list each at keys 'x', 'y', 'w', 'h'.
        :param coords: The coords that get added. A dictionary with one value each at keys 'x', 'y', 'w', 'h'.
        :return: The dictionary containing the added coords.
    '''

    dict['x'].append(coords['x'])
    dict['y'].append(coords['y'])
    dict['w'].append(coords['w'])
    dict['h'].append(coords['h'])

    return dict

def detect_page(region):
    '''
        Detects whether the corresponding page is on left or right side of a book based on the passed coords of a page number region. Conditions for this decision are defined globally (NUMBERREGION_THRESHOLD_MAX_W_COORD, NUMBERREGION_THRESHOLD_MAX_X_COORD and NUMBERREGION_THRESHOLD_MIN_X_COORD)

        :param region_coords: A dictionary containing integers on keys 'x', 'y', 'w', 'h'
        :return: Returns either 'l', 'r' or 'not a pageNr'
    '''
    region_coords = region.coords.box
    if region_coords['w'] < NUMBERREGION_THRESHOLD_MAX_W_COORD and region_coords['x'] < NUMBERREGION_THRESHOLD_MAX_X_COORD:                
        # Left page
        return 'l', region.lines[0].text
    elif region_coords['w'] < NUMBERREGION_THRESHOLD_MAX_W_COORD and region_coords['x'] > NUMBERREGION_THRESHOLD_MIN_X_COORD:                
        # Right page
        return 'r', region.lines[0].text
    else:
        return 'not a pageNr', ''
            
def classify_left_page_columns(scan, textRegion_dict, unknown_region_dict, final_regest_dict):
    '''
        Iterates through each region of the passed scan and tries to classify it based on the globally defined thresholds for left page. Page Number region doesnt get recognised because its only relevant for page side detection.
        Also makes basic classification of all lines from the table region filling the textRegion_dict with the coordinates of each line, its text and a basic classification of either on of the three columns ('date', 'place', 'regest_text')

        :param scan: A PageXMLScan object from the physical_document_model containing all relevant pagexml informations.
        :param textRegion_dict: A dictionary containing informations of the lines of the table region. Must contain lists on keys 'x', 'y', 'w', 'h', 'type', 'text'.
        :param unknown_region_dict: The dictioniary regions get added to if they dont fit in the conditions of all other regions Must contain lists on keys 'x', 'y', 'w', 'h'.
        :return: textRegion_dict, unknown_region_dict
    '''

    for region in scan.text_regions:
        if len(region.lines) > 0: # If region doesnt have lines its not relevant
            # Iterate through each text region and try to classify it based on coordinates and the global defined thresholds
            region_coords = region.coords.box

            if region_coords['w'] > TEXTREGION_THRESHOLD_MIN_WIDTH and region_coords['h'] > TEXTREGION_THRESHOLD_MIN_HEIGHT:
                # Table region
                for line in region.lines:
                    # Iterate through each line of the table region and try to classify it based coordinates and the global defined thresholds         
                    line_coords = line.coords.box
                    add_coords_to_dict(textRegion_dict, line_coords)

                    if line_coords['x'] > REGEST_THRESHOLD_MIN_X_COORD_LEFT:  
                        # Regest text column
                        textRegion_dict['type'].append('regest_text')
                    elif line_coords['x'] < REGEST_THRESHOLD_MIN_X_COORD_LEFT and line_coords['x'] > PLACE_THRESHOLD_MIN_X_COORD_LEFT:
                        # Regest place column
                        textRegion_dict['type'].append('place')
                    else:
                        # Regest date column
                        textRegion_dict['type'].append('date')

                    textRegion_dict['text'].append(line.text)                 

            elif region_coords['x'] > HEADERREGION_THRESHOLD_MIN_X_COORD and region_coords['x'] < HEADERREGION_THRESHOLD_MAX_X_COORD and region_coords['h'] < HEADERREGION_THRESHOLD_MAX_H_COORD:
                # Header region
                pope = ''
                for char in region.lines[0].text:
                    if char != '.':
                        pope += char
                    else:
                        pope += char
                        break
                final_regest_dict['pope'].append(pope)
            else:
                # Unknown region
                add_coords_to_dict(unknown_region_dict, region_coords)

    return textRegion_dict, unknown_region_dict, final_regest_dict

def classify_right_page_columns(scan, textRegion_dict, unknown_region_dict, final_regest_dict):
    '''
        Iterates through each region of the passed scan and tries to classify it based on the globally defined thresholds for right page. Page Number region doesnt get recognised because its only relevant for page side detection.
        Also makes basic classification of all lines from the table region filling the textRegion_dict with the coordinates of each line, its text and a basic classification of either on of the three columns ('date', 'place', 'regest_text')

        :param scan: A PageXMLScan object from the physical_document_model containing all relevant pagexml informations.
        :param textRegion_dict: A dictionary containing informations of the lines of the table region. Must contain lists on keys 'x', 'y', 'w', 'h', 'type', 'text'.
        :param unknown_region_dict: The dictioniary regions get added to if they dont fit in the conditions of all other regions Must contain lists on keys 'x', 'y', 'w', 'h'.
        :return: textRegion_dict, unknown_region_dict
    '''

    for region in scan.text_regions:
        if len(region.lines) > 0: # If region doesnt have lines its not relevant
            # Iterate through each text region and try to classify it based on coordinates and the global defined thresholds
            region_coords = region.coords.box

            if region_coords['w'] > TEXTREGION_THRESHOLD_MIN_WIDTH and region_coords['h'] > TEXTREGION_THRESHOLD_MIN_HEIGHT:
                # Table region
                for line in region.lines:
                    # Iterate through each line of the table region and try to classify it based coordinates and the global defined thresholds         
                    line_coords = line.coords.box
                    add_coords_to_dict(textRegion_dict, line_coords)

                    if line_coords['x'] > REGEST_THRESHOLD_MIN_X_COORD_RIGHT:  
                        # Regest text column
                        textRegion_dict['type'].append('regest_text')
                    elif line_coords['x'] < REGEST_THRESHOLD_MIN_X_COORD_RIGHT and line_coords['x'] > PLACE_THRESHOLD_MIN_X_COORD_RIGHT:
                        # Regest place column
                        textRegion_dict['type'].append('place')  
                    else:
                        # Regest date column
                        textRegion_dict['type'].append('date')

                    textRegion_dict['text'].append(line.text)                 

            elif region_coords['x'] > HEADERREGION_THRESHOLD_MIN_X_COORD and region_coords['x'] < HEADERREGION_THRESHOLD_MAX_X_COORD and region_coords['h'] < HEADERREGION_THRESHOLD_MAX_H_COORD:
                # Header region
                pope = ''
                for char in region.lines[0].text:
                    if char != '.':
                        pope += char
                    else:
                        pope += char
                        break
                final_regest_dict['pope'].append(pope)
            else:
                # Unknown region
                add_coords_to_dict(unknown_region_dict, region_coords)

    return textRegion_dict, unknown_region_dict, final_regest_dict

def classify_regest_start(df):
    '''
        Classifies all indented lines (regest_start) of the third column (regest_text) based on the lines on the far left and the mean starting coordinate of all lines.

        :param: df: A Pandas DataFrame created from the textRegion_dict structure containing all lines of the table region.
        :return: df
    '''

    regest_df = df[(df['type'] == 'regest_text')] # Only get lines of third column
    x_min = regest_df['x'].min()
    x_avg = regest_df['x'].mean()
    x_max = regest_df['x'].max()
    regestStart_threshold_max_x_Coord = x_min + x_avg*0.07 # The rows on the far left, including some space if the rows as a whole are moved slightly to the right over time
    for idx, row in df.iterrows():
        if row['type'] == 'regest_text' and row['x'] < regestStart_threshold_max_x_Coord and row['x'] != x_max:            
            df.at[idx, 'type'] = 'regest_start'

    return df

def classify_regest_date(df):
    '''
        Classifies all years headers that are part of the third columns of the indented lines (regest_start) based on the lines on the far left and the mean starting coordinate of all indented lines since they are even more far left than the regest start lines.

        :param: df: A Pandas DataFrame created from the textRegion_dict structure containing all lines of the table region.
        :return: df
    '''

    regest_df = df[(df['type'] == 'regest_start')] # Only get rows tha are indented
    x_min = regest_df['x'].min()
    x_avg = regest_df['x'].mean()
    w_avg = regest_df['w'].mean()
    h_max = regest_df['h'].max()
    h_avg = regest_df['h'].mean()
    regestDate_threshold_max_x_Coord = x_min + x_avg*0.015
    regestDate_threshold_max_w_Coord = x_avg + w_avg - w_avg*0.025
    regestDate_threshold_min_h_Coord = h_max - h_avg*0.075
    for idx, row in df.iterrows():
        if row['type'] == 'regest_start' and row['x'] < regestDate_threshold_max_x_Coord and row['text'][0] in ['0', '1', '2', '3', '4', '5', '6', '7', '8', '9'] : # The rows on the far left, including some space if the rows as a whole are moved slightly to the right over time
            if row['w'] < regestDate_threshold_max_w_Coord or row['h'] > regestDate_threshold_min_h_Coord: # The rows dont spread over the whole distance and are a little bit higher than other indented lines
                df.at[idx, 'type'] = 'regest_date'

    return df

def detect_new_pope(df):
    '''
        Classifies all pope headers and the following pope unnecessary overview that are part of the third columns based on the their position in the middle of the column and the fact thath there is no date in the next line (in this case a new regest would start afterwards). 

        :param: df: A Pandas DataFrame created from the textRegion_dict structure containing all lines of the table region.
        :return: The DataFrame without pope stuff
    '''
    new_pope_idx = []
    new_pope = False
    is_next_date = False
    df_text = df[(df['type'] == 'regest_text')]
    x_min = df_text['x'].min()
    x_avg = df_text['x'].mean()    
    regestDate_threshold_max_x_Coord = x_min + x_avg*0.015    
    for idx, row in df.iterrows():
        is_next_date = False
        if idx > 0 and idx < len(df.index)-1:
            if df.iloc[idx+1]['x'] < regestDate_threshold_max_x_Coord:
                is_next_date = True
            if new_pope == False:
                y_distance = row['y'] - df.iloc[idx-1]['y']
                if y_distance >= NEW_POPE_THRESHOLD_Y_COOD and is_next_date == False: # Current Row is a pope header and the next row isnt a date: Begin of unnecessary pope lines
                    new_pope_idx.append(idx)
                    new_pope = True
                elif y_distance >= NEW_POPE_THRESHOLD_Y_COOD and is_next_date == True: # Pope header but next row is date: standard regests begin again
                    new_pope_idx.append(idx)
            else:
                # Classify as unnecessary pope line as long as no new date line begins
                if is_next_date == True:
                    new_pope = False
                new_pope_idx.append(idx)
        elif idx == len(df.index)-1 and new_pope == True:
            new_pope_idx.append(idx)   
    df = df.drop(index=new_pope_idx)
    df = df.reset_index(drop=True)

    return df

def detect_pope_overview(df):
    '''
    Detects whether the page started with a pope overview start that needs to get dropped and didnt get recognised in detect_new_pope()

    :param: df: The dataframe to check
    :return: The Pandas DataFrame without pope overview
    '''
    drop_idx = []
    new_pope_overview = False
    df_text = df[(df['type'] == 'regest_text')]
    x_min = df_text['x'].min()
    x_avg = df_text['x'].mean()    
    regestDate_threshold_max_x_Coord = x_min + x_avg*0.015
    for idx, row in df.iterrows():
        if idx == 0:
            if row['type'] == 'regest_text': # Normal regest pages always start with a date. Otherwise ist part of pope overview that startet on previous page, so it didnt get recognised in detect_new_pope()
                new_pope_overview = True
        if df.iloc[idx]['x'] >= regestDate_threshold_max_x_Coord and new_pope_overview == True:
            drop_idx.append(idx)
        elif df.iloc[idx]['x'] < regestDate_threshold_max_x_Coord and new_pope_overview == True:
            new_pope_overview = False
    df = df.drop(index=drop_idx)
    df = df.reset_index(drop=True)
    
    return df

def get_closest_regest(df, df_regest_start, type):
    '''
       Creates a list containing the idx of each row in df and idx of the closest regest starting line  to that row (of df). If two rows have the same closest starting regest, pop the last one as it is most propablly something unnecessary in that case.

       :param: df: The Pandas DataFrame containing the rows to which we want to find the closest regests.
       :param: df_regest_start: A Pandas DataFrame containing only thr rows that are the start of a new regest.
       :param: type:
       :return: An list of indices containing tuples of row idx, closest regest idx and the distance
    '''
    closest_regest = []
    for idx, line in df.iterrows():
        min_distance = 9999999
        min_regest_idx = 0
        for regest_idx, regest_line in df_regest_start.iterrows():
            if abs(regest_line['y'] - line['y']) < min_distance:
                min_distance = abs(regest_line['y'] - line['y'])
                min_regest_idx = regest_idx
        if len(closest_regest) > 0:
            if closest_regest[-1][1] == min_regest_idx and type == 'date': # Many pages have an unnecessary text at the bottom left that gets classified as date. So if the last two date have the same closest neighbour, remove the last one as it is most propably the unnecassary text              
                if closest_regest[-1][2] >= min_distance:
                    closest_regest.pop()
                    closest_regest.append([idx, min_regest_idx, min_distance])
            else:
                closest_regest.append([idx, min_regest_idx, min_distance])
        else:
            closest_regest.append([idx, min_regest_idx, min_distance])
    return closest_regest

def get_regest_number(txt):
    '''
    Slices the given string so that the regest number at the start of the string gets seperated.

    :param: txt: The string to extract the regest number from
    :return: The seperated regest number and the other text
    '''

    nr_chars = ['0', '1', '2', '3', '4', '5', '6', '7', '8', '9', '(', ' ']
    slice_idx = 0
    if txt is not None:
        for idx, char in enumerate(txt):
            if char in nr_chars:
                pass
            else:
                if idx == 0:
                    if len(txt) > 1:
                        if txt[idx+1] not in nr_chars:
                            slice_idx = idx
                            break
                else:
                    if txt[idx-1] == ' ':
                        slice_idx = idx
                    elif txt[idx-1] == '(':
                        slice_idx = idx-1
                    else:
                        slice_idx = idx+1
                    break
        regest_nr = txt[:slice_idx]
        txt = txt[slice_idx:]
        if len(txt) > 0:
            if txt[0] == ' ':
                txt = txt[1:]
        
        if len(regest_nr) > 1:
            if regest_nr[-1] == ' ':
                regest_nr = regest_nr[:-1]

    else:
        regest_nr = False

    return regest_nr, txt

def get_year(txt):
    '''
    Slices the given string so that the year at the start of the string gets seperated.

    :param: txt: The string to extract the year from
    :return: The seperated year
    '''

    year_chars = ['0', '1', '2', '3', '4', '5', '6', '7', '8', '9', '?', '—', '-', ' ']
    slice_idx = len(txt)
    for idx, char in enumerate(txt):
        if char in year_chars:
            pass
        else:
            slice_idx = idx
            break
    txt = txt[:slice_idx]

    return txt

def classify_whole_regests(df, final_regest_dict):
    '''
        Gets a Pandas DataFrame whichs lines already got classified and tries to detect which lines make one regest.

        :param: df: The pandas DataFrame containing the lines to classify
        :param: final_regest_dict: A dictionary conatining lists on keys date, place, number and text to which the final classified regests get added
        :return: The passed DataFrame and the passed dictionary filled with all regests
    '''
    current_year = ''
    current_place = ''
    regest_full_text = []
    first_regest = True
    df_regest_start = df[(df['type'] == 'regest_start')]
    df_date = df[(df['type'] == 'date')]
    df_place = df[(df['type'] == 'place')]
    date_closest_regest = get_closest_regest(df_date, df_regest_start, 'date')
    place_closest_regest = get_closest_regest(df_place, df_regest_start, 'place')
   
    for idx, row in df.iterrows():
        if row['type'] == 'regest_date': # Get the current year as it is'nt a part of the date column.
            current_year = get_year(row['text'])
        elif row['type'] == 'regest_start':
            # new regest
            if first_regest == False:
                regest_full_text = [i for i in regest_full_text if i is not None]
                regest_full_text = ' '.join(regest_full_text)
                final_regest_dict['text'].append(regest_full_text)
                regest_nr, txt = get_regest_number(row['text'])
            elif first_regest == True and len(regest_full_text) > 0: # Regest started on previous page.
                final_regest_dict['date'] = ['prev_page']
                final_regest_dict['place'] = ['prev_page']
                final_regest_dict['number'] = ['prev_page']
                regest_full_text = [i for i in regest_full_text if i is not None]
                regest_full_text = ' '.join(regest_full_text)
                final_regest_dict['text'].append(regest_full_text)
                txt = row['text']
                regest_nr == False # No regest number                
                first_regest = False
            else:
                regest_nr, txt = get_regest_number(row['text'])
                first_regest = False

            regest_full_text = []
            if txt is not None:
                regest_full_text.append(txt)

            # Regest number
            if regest_nr != False:
                final_regest_dict['number'].append(regest_nr)
            else: 
                final_regest_dict['number'].append(' ')

            # Get the date of regest
            full_date = current_year
            for i in date_closest_regest:
                if i[1] == idx:
                    if df_date.loc[i[0]]['text'] is not None:
                        full_date += ' ' + df_date.loc[i[0]]['text']
                    else:
                        full_date += ' '
            final_regest_dict['date'].append(full_date)

            # Get the place of regest
            place = ' '
            for i in place_closest_regest:
                if i[1] == idx:
                    if df_place.loc[i[0]]['text'] is not None:
                        place = df_place.loc[i[0]]['text']
            if not any(char.isalpha() for char in place):
                if current_place != '':
                    if '(„)' in place or '(„' in place or '„)' in place:
                        place = '(' + current_place + ')'
                        final_regest_dict['place'].append(place)
                    elif '„' in place:
                        final_regest_dict['place'].append(current_place)
                    else:
                        final_regest_dict['place'].append(' ')
                else:
                    if '(„)' in place or '(„' in place or '„)' in place:
                        place = '(' + current_place + ')'
                        final_regest_dict['place'].append('(prev_regest)')
                    elif '„' in place:
                        final_regest_dict['place'].append('prev_regest')
                    else:
                        final_regest_dict['place'].append(' ')
            else:
                final_regest_dict['place'].append(place)
                current_place = place

        elif row['type'] == 'regest_text' and first_regest == False:
            if row['text'] is not None:
                regest_full_text.append(row['text'])
        elif row['type'] == 'regest_text' and first_regest == True:
            final_regest_dict['date'] = ['prev_page']
            final_regest_dict['place'] = ['prev_page']
            final_regest_dict['number'] = ['prev_page']
            if row['text'] is not None:
                regest_full_text.append(row['text'])
            first_regest = False

    if len(regest_full_text) > 0:
        regest_full_text = [i for i in regest_full_text if i is not None]
        regest_full_text = ' '.join(regest_full_text)
        final_regest_dict['text'].append(regest_full_text)

    return df, final_regest_dict

def drop_unnecessary(df):
    '''
        Detects lines that are unnecessary for regest classification and dont contain necessary information.

        :param: df: The Pandas DataFrame to drop the unnecessary lines from
        :return: The passed DataFrame without unnecessary lines
    '''
    drop_idx = []
    for idx, row in df.iterrows():
        if row['type'] == 'date' and row['h'] < 130 and row['w'] > 500: # Sometimes there is a text at the bottom of date column which doesnt contain relevant information.
            drop_idx.append(idx)
        if row['type'] == 'regest_text' and row['x'] > 2300 and row['y'] > 3450 and row['w'] < 100: # Sometimes there are numbers at the bottom right corner that are not needed.
            drop_idx.append(idx)
    df = df.drop(index=drop_idx)
    df = df.reset_index(drop=True)

    return df

def classify(scan):
    '''
        Main classification function. Iterates through each region of the passed scan and manages classification of table region 

        :param: scan: The PageXML scan to classify the regests from
        :return: A Pandas DataFrame containing positional data for plotting and Pandas Dataframe containing all regests
    '''
    textRegion_dict, unknown_region_dict, final_regest_dict = new_dictionaries()
    for region in scan.text_regions:
        if len(region.lines) > 0:                
            page, page_nr = detect_page(region)
            if page == 'l':
                final_regest_dict['page_nr'].append(page_nr)
                textRegion_dict, unknown_region_dict, final_regest_dict = classify_left_page_columns(scan, textRegion_dict, unknown_region_dict, final_regest_dict)
                break
            elif page == 'r':
                final_regest_dict['page_nr'].append(page_nr)
                textRegion_dict, unknown_region_dict, final_regest_dict = classify_right_page_columns(scan, textRegion_dict, unknown_region_dict, final_regest_dict)
                break
    df = pd.DataFrame(textRegion_dict)
    df = df.sort_values('y')
    df = df.reset_index(drop=True)
    df = detect_new_pope(df)
    df = detect_pope_overview(df)
    df = drop_unnecessary(df)
    df = classify_regest_start(df)
    df = classify_regest_date(df)
    
    df, final_regest_dict = classify_whole_regests(df, final_regest_dict)

    i = 0
    while i < len(final_regest_dict['date']):
        if i >= len(final_regest_dict['pope']):
            if len(final_regest_dict['pope']) == 0:
                final_regest_dict['pope'].append('')
            else:
                final_regest_dict['pope'].append(final_regest_dict['pope'][0])
        if i >= len(final_regest_dict['page_nr']):
            if len(final_regest_dict['page_nr']) == 0:
                final_regest_dict['page_nr'].append('')
            else:
                final_regest_dict['page_nr'].append(final_regest_dict['page_nr'][0])
        i += 1

    print(len(final_regest_dict['pope']), len(final_regest_dict['date']), len(final_regest_dict['place']), len(final_regest_dict['number']), len(final_regest_dict['text']), len(final_regest_dict['page_nr']))
    final_df = pd.DataFrame(final_regest_dict)

    return df, final_df

def plot_text_regions(df):
    '''
        Creates a window using matplotlib that plots the regions of all rows.
        :param: df: The pandas DataFrame to plot the image from
    '''
    fig, ax = plt.subplots(figsize=(10, 8))
    for idx, row in df.iterrows():
        if row['type'] == 'date':
            color = 'g'         
        elif row['type'] == 'place':
            color = 'b'
        elif row['type'] == 'regest_text':
            color = 'y'
        elif row['type'] == 'regest_start':
            color = 'r'        
        elif row['type'] == 'regest_date':
            color = 'c'          
        elif row['type'] == 'new_pope':
            color = 'm'             
        else:
            color = 'k'
        rect = patches.Rectangle((row['x'], row['y']), row['w'], row['h'], linewidth=1, edgecolor=color, facecolor='none')
        ax.add_patch(rect)
        #ax.text(row['x'], row['y'], row['text'], fontsize=8, verticalalignment='top')
    ax.set_xlim(0, df['x'].max() + df['w'].max())
    ax.set_ylim(df['y'].max() + df['h'].max(), 0)
    plt.axis('off')
    plt.show()

def ask_for_plot():
    print('Do you want to plot the line regions of a specific file? (y/n)')
    yes = ['y', 'Y', 'yes', 'Yes']
    no = ['n', 'N', 'no', 'No']
    i = input()
    if i in yes:
        print('----------------------------------------------------------------------------------------------------------------------')
        return True
    elif i in no:
        print('----------------------------------------------------------------------------------------------------------------------')
        return False
    else:
        print('Invalid input. Continuing without plotting...')
        return False

def combine_regests(df):
    '''
        Deletes all lines categorised as belonging to the previous page and adds their text to previous line text. Also detects all regest thats places are categorised to be the same as previous places regest
    
        :param: df: The Pandas DataFrame whichs regests get merged
        :return: The Pandas DataFrame without rows that belong to previous regest
    '''
    # Merge text
    drop_idx = []
    for idx, row in df.iterrows():
        if row['date'] == 'prev_page' and row['place'] == 'prev_page' and row['number'] == 'prev_page':
            if idx != 0:
                txt = ' '.join([df.iloc[idx-1]['text'], row['text']])
                df.loc[idx-1, 'text'] = txt
                drop_idx.append(idx)
            else:
                drop_idx.append(idx)
    df = df.drop(index=drop_idx)
    df = df.reset_index(drop=True)    

    # Detect places that are classified as being the same as previous place
    for idx, row in df.iterrows():
        if row['place'] == 'prev_regest':
            i = idx-1
            while i > 0:
                if any(char.isalpha() for char in df.loc[i, 'place']):
                    weg = ['(', ')']
                    place = ''.join([p for p in df.loc[i, 'place'] if p not in weg])
                    df.loc[idx, 'place'] = place
                    i = 0
                i = i-1
        elif row['place'] == '(prev_regest)':
            i = idx-1
            while i > 0:
                if any(char.isalpha() for char in df.loc[i, 'place']):
                    df.loc[idx, 'place'] = '(' + df.loc[i, 'place'] + ')'
                    i = 0
                i = i-1                

    return df

def output(df, format, file_name):
    '''
        Exports the given pandas DataFrame into csv, xml or tsv file depending on passed format string.

        :param: df: The Pandas DataFrame to export
        :param: format: Either 'csv' or 'tsv' depending on favored output format.    
    '''
    file_name = file_name[:-4]
    if format.lower() == 'csv':
        df.to_csv(OUTPUT_PATH + '\\' + file_name + 'CSV.csv', index=True)
    elif format.lower() == 'tsv':
        df.to_csv(OUTPUT_PATH + '\\' + file_name + 'TSV.tsv', index=True, sep='\t')
    elif format.lower() == 'single_xml':
        df.to_xml(OUTPUT_PATH + '\\' + file_name + 'XML.xml')
    elif format.lower() == 'multi_xml':
        zip = OUTPUT_PATH + '\\' + file_name + '.zip'
        with ZipFile(zip, 'w') as zipFile:
            for i in range(len(df)):
                single_regest = df.iloc[[i]]
                with io.BytesIO() as puffer:
                    single_regest.to_xml(puffer, index=False, encoding='utf-8')
                    puffer.seek(0)
                    zipFile.writestr(f"{single_regest['number'].values[0]}.xml", puffer.read())
    else:
        print('Invalid output format. COuld not create output file.')

# ------------------------------------------------------------------------------------
# ------------------------------------------------------------------------------------
# Main
# ------------------------------------------------------------------------------------
# ------------------------------------------------------------------------------------
input_type = console_welcome()

# Single page classification
if input_type == 'single':
    # Enter file name
    print('Please enter file name (Must be in PageXML folder and end in .xml)')
    file_name = input()
    single_file = INPUT_PATH + '\\' + file_name
    if os.path.exists(single_file) != True:
        print('Path doesnt exist. Stopping process...')
        quit()
    print('----------------------------------------------------------------------------------------------------------------------')

    scan = pxml_parser.parse_pagexml_file(single_file)
    df, final_df = classify(scan)
    #print(final_df)

    # Plot the regions if wanted
    if ask_for_plot() == True:
        plot_text_regions(df)

    # Ask for output format
    print('----------------------------------------------------------------------------------------------------------------------')
    print('Please enter output format (csv/tsv//single_xml/multi_xml)')
    output_format = input()

    output(final_df, output_format, file_name)

# Zip archive classification
elif input_type == 'zip':
    # Enter file name
    print('Please enter zip archive name (Must end in .zip)')
    file_name = input()
    zip_file = INPUT_PATH + '\\' + file_name
    if os.path.exists(zip_file) != True:
        print('Path doesnt exist. Stopping process...')
        quit()
    print('----------------------------------------------------------------------------------------------------------------------')

    # Ask if a specific page should get plotted
    plot_file = ''
    if ask_for_plot() == True:
        print('Please enter file name of the page you want to plot (Must be inpage the zip archive and end in .xml)')
        plot_file = input()

    # Ask for output format
    print('----------------------------------------------------------------------------------------------------------------------')
    print('Please enter output format (csv/tsv//single_xml/multi_xml)')
    output_format = input()

    print('Processing...')
    zip_df = []
    for info, data in read_page_archive_files(zip_file):
        if info['archived_filename'] != 'METS.xml':
            scan = pxml_parser.parse_pagexml_file(pagexml_file=info['archived_filename'], pagexml_data=data)
            df, final_df = classify(scan)
            zip_df.append(final_df)

            # Plot specific page if wanted
            if plot_file != '' and plot_file == info['archived_filename']:
                plot_text_regions(df)

    # Merge all dataframe together
    zip_df = pd.concat(zip_df, ignore_index=True)
    zip_df = combine_regests(zip_df)

    output(zip_df, output_format, file_name)
    print('Finished :)')

# Invalid Input
else:
    print('Invalid input. Stopping process...')

'''
Noch zu tun:

- Regesten ohne Nummer -> erweiterte Nummer der vorherigen Regeste z.B: 10614-02 oder:
- Seitennummer hinzufügen als eigenes XML-Element, da wichtig für Regesten ohne Jaffé-Nummer
- Papst bzw. Aussteller als XML Element
- Generell Regionenerkennung verbessern
- Bei Datum Anführungszeichen ersetzen
- Recipit
- Aufräumen :)

'''