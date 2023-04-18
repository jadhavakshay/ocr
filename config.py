# folder path
input_dir_path = r'D:\Spherex\RQ_edit_6\input'
output_dir_path = r'D:\Spherex\RQ_edit_6\output'

env = 'QA'

if env == 'DEV':
    # DEV
    DB_HOST = 'spherexratings-local.mysql.database.azure.com'
    DB_NAME = 'questionnaire_ratings'
    DB_USER_NAME = 'dev_rule_engine@spherexratings-local'
    DB_PASSWORD = 'A_AW],rbkiygeOC~'
elif env == 'QA':
    # QA
    DB_HOST = 'spherexratings-local.mysql.database.azure.com'
    DB_NAME = 'questionnaire_ratings_qa'
    DB_USER_NAME = 'qa_rule_engine@spherexratings-local'
    DB_PASSWORD = "B0XS5,Oyi>BrK,h)"
elif env == 'UAT':
    # UAT
    DB_HOST = 'spherexratings-dev.mysql.database.azure.com'
    DB_USER_NAME = 'uat_questionnaire@spherexratings-dev'
    DB_PASSWORD = 'PtufoTtGUeIHXhR'
    DB_NAME = "questionnaire_ratings_uat"


CLIENT = env

TITLE_TYPE = 1
FEATURE_NAME = 2
SERIES_NAME = 3
SERIES_NUMBER = 4
EPISODE_NUMBER = 5
EPISODE_NAME = 6
VERSION = 7
EXT_CONTENT_ID = 8
IMDB_ID = 9
COUNTRY_OF_ORIGIN = 10
ORIGINAL_LANGUAGE = 11
RUNTIME = 12
RELEASE_YEAR = 13
SYNOPSIS = 14
WRITER = 15
DIRECTOR = 16
PRODUCER = 17
PRODUCTION_COMPANY = 18
CAST = 19
US_RATING = 20
US_RATING_ADVISORY = 21
ARTWORK = 22
ORIGINAL_AIR_DATE = 23
TERRITORIES = 24

if CLIENT == 'DEV':
    QUESTIONNAIRE_QUESTION_ID = 1
    QUESTIONNAIRE_OPTION_ID = 2
    QUESTIONNAIRE_SELECTED_ANSWER = 15

    PROFANITY_QUESTION_ID = 1
    PROFANITY_OPTION_ID1 = 3
    PROFANITY_OPTION_ID2 = 5
    PROFANITY_OPTION_ID3 = 7
    PROFANITY_SELECTED_ANSWER = 9
else:
    QUESTIONNAIRE_QUESTION_ID = 2
    QUESTIONNAIRE_OPTION_ID = 3
    QUESTIONNAIRE_SELECTED_ANSWER = 17

    PROFANITY_QUESTION_ID = 2
    PROFANITY_OPTION_ID1 = 4
    PROFANITY_OPTION_ID2 = 6
    PROFANITY_OPTION_ID3 = 8
    PROFANITY_SELECTED_ANSWER = 14

