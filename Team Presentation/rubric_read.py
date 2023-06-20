import os
from skimage import io
import pytesseract
import numpy as np

import pandas as pd
from pypdf import PdfMerger

import matplotlib.pyplot as plt


pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'


# Return page pixel given inch dimensions and 300dpi
#-------------------------------------------------------------------------
def in2px(inches):
    return int(300*inches)

# scan horizontal text based on y-position as bottom of text
#-------------------------------------------------------------------------
def scan_text(image, v_position):

    x1 = in2px(1)
    x2 = in2px(7.5)
    y1 = in2px(v_position-0.2)
    y2 = in2px(v_position+0.1)

    ocr_text = pytesseract.image_to_string( image[y1:y2, x1:x2] )
    ocr_text = ocr_text.replace('\n', '')
    ocr_text = ocr_text.replace(':', ' -')
    return ocr_text

# scan horizontal bubbles based on y-position as bottom of bubbles and text
#-------------------------------------------------------------------------
def scan_bubbles(image, v_position):
    
    r = 0.1        # radius of bubble in inches
    low = 255      # lowest average intensity of bubble
    selection = 0  # most filled selection (1-4)

    # iterate through possible bubbles
    for i in range(4):
        x = 0.5*i+4.75

        # Circle center is drawn at
        # x =     in2px(0.5*i+4.75)
        # y =     in2px(v_position)-r

        x1 = in2px(x-r/2)
        x2 = in2px(x+r/2)
        y1 = in2px(v_position-3*r/2) # up page 3/2r
        y2 = in2px(v_position-r/2)   # up page 1/2r

        avg = np.mean(image[y1:y2, x1:x2])
        
        if (avg < low) and (avg < 200):
            selection = i+1
            low = avg
    
    return selection

# Get team comments
#-------------------------------------------------------------------------
def scan_team_comments(image):
    x1 = in2px(1)
    x2 = in2px(7.5)
    y1 = in2px(3.875)
    y2 = in2px(3.875+4)
    
    return image[y1:y2, x1:x2]

# Get individual comments
#-------------------------------------------------------------------------
def scan_individual_comments(image, v_position):
    x1 = in2px(1)
    x2 = in2px(7.5)
    y1 = in2px(v_position)
    y2 = in2px(v_position+1.4)
    
    return image[y1:y2, x1:x2]

# Get team data
#-------------------------------------------------------------------------
def get_team_results(image, team_criteria, team_ratings):

    # initialize team results
    team_results = {}
    
    # determine team name
    team_name = scan_text(image, 1.125)
    # team_name = team_name[6:-1]
    team_results['Team'] = team_name

    # determine criteria scores
    y = 1.625
    for criteria in team_criteria:
        selection = scan_bubbles(image, y)

        if selection:
            team_results[criteria] = team_ratings[selection-1]

        y += 0.25

    return pd.Series(team_results)

# Get individual data
#-------------------------------------------------------------------------
def get_individual_results(image, individual_criteria, individual_ratings, v_position):

    # initialize individual results
    individual_results = {}
    
    # determine individual name
    individual_name = scan_text(image, v_position)
    # individual_name = individual_name[:-1]
    individual_results['Name'] = individual_name

    # determine criteria scores
    y = v_position+0.5
    for criteria in individual_criteria:
        selection = scan_bubbles(image, y)

        if selection:
            individual_results[criteria] = individual_ratings[selection-1]

        y += 0.25

    return pd.Series(individual_results)

#-------------------------------------------------------------------------
# Load and Organize Rubric Criteria from CSV file
#-------------------------------------------------------------------------
path = os.path.dirname(__file__)
raw_data = pd.read_csv(path+'/rubric_criteria.csv')

team_criteria = list(raw_data['Team Criteria'].dropna())
team_ratings = list(raw_data['Team Ratings'].dropna())
individual_criteria = list(raw_data['Individual Criteria'].dropna())
individual_ratings = list(raw_data['Individual Ratings'].dropna())


#-------------------------------------------------------------------------
# Read and return the results for all the teams in a directory
#-------------------------------------------------------------------------

# location of scanned images
path = os.path.dirname(__file__)
filenames = ['scan_ahickers_2023-06-20-14-08-43_1.jpeg',
             'scan_ahickers_2023-06-20-14-08-43_4.jpeg']

team_results = pd.DataFrame()

i = 1
for filename in filenames:
    if filename.endswith(".jpeg") or filename.endswith(".jpg"):
        # import scanned image
        scan = io.imread(path + '/scans/' + filename)

        # Get team data
        current_team_results = get_team_results(scan, team_criteria, team_ratings)
        team_results = team_results._append(current_team_results, ignore_index=True)

        # save iamge of comments
        io.imsave(path + '/scans/' + current_team_results['Team'] + ' - comment_' + str(i) + '.pdf', scan_team_comments(scan))
        i += 1

team_results.to_csv(path + '/results/' + 'team results.csv')


#-------------------------------------------------------------------------
# Read and return the results of individuals for all the scans in a directory
#-------------------------------------------------------------------------

# location of scanned images
path = os.path.dirname(__file__)
filenames = ['scan_ahickers_2023-06-20-14-08-43_2.jpeg',
             'scan_ahickers_2023-06-20-14-08-43_3.jpeg']

ind_results = pd.DataFrame()

i = 1
for filename in filenames:
    if filename.endswith(".jpeg") or filename.endswith(".jpg"):
        # import scanned image
        scan = io.imread(path + '/scans/' + filename)

        # Get individual data
        for i in range(3):
            current_ind_results = get_individual_results(scan, individual_criteria, individual_ratings, i*3+1.125)
            ind_results = ind_results._append(current_ind_results, ignore_index=True)

            # save image of comments
            io.imsave(path + '/scans/' + current_ind_results['Name'] + ' - comment_' + str(i) + '.pdf', scan_individual_comments(scan, i*3+2.375))
        i += 1

ind_results.to_csv(path + '/results/' + 'rubric individual results.csv')





#-------------------------------------------------------------------------
# Analyze and plot results
#-------------------------------------------------------------------------

teams = list(team_results['Team'].unique())

for team in teams:
    # filter results by team
    team_summary =  team_results[team_results['Team'] == team]

    # get value counts for each team
    result_counts = team_summary[team_criteria].apply(pd.value_counts)
    result_counts = result_counts.fillna(0)
    result_counts = result_counts.transpose()

    # plot value counts and save
    result_counts.plot(kind='barh', stacked=True)
    plt.title(team)
    plt.tight_layout()
    plt.savefig(path + '/scans/' + team + ' - barplot.pdf')

individuals = list(ind_results['Name'].unique())

for individual in individuals:
    # filter results by team
    individual_summary =  ind_results[ind_results['Name'] == individual]

    # get value counts for each team
    result_counts = individual_summary[individual_criteria].apply(pd.value_counts)
    result_counts = result_counts.fillna(0)
    result_counts = result_counts.transpose()

    # plot value counts and save
    result_counts.plot(kind='barh', stacked=True)
    plt.title(individual)
    plt.tight_layout()
    plt.savefig(path + '/scans/' + individual + ' - barplot.pdf')

#-------------------------------------------------------------------------
# Combine into PDF
#-------------------------------------------------------------------------

for team in teams:

    merger = PdfMerger()
    
    for filename in os.listdir(path + '/scans/'):
        
        if team in filename:
            merger.append(path + '/scans/' + filename)

    merger.write(path + '/results/' + team +'.pdf')
    merger.close()


for individual in individuals:

    merger = PdfMerger()
    
    for filename in os.listdir(path + '/scans/'):
        
        if individual in filename:
            merger.append(path + '/scans/' + filename)

    merger.write(path + '/results/' + individual +'.pdf')
    merger.close()