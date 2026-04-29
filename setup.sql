-- Run this script once to deploy the Cortex chatbot as a Streamlit in Snowflake app.
-- Adjust the database, schema, warehouse, and stage names to match your environment.

USE ROLE SYSADMIN;

-- Create a dedicated database/schema if needed
CREATE DATABASE IF NOT EXISTS CORTEX_APPS;
CREATE SCHEMA  IF NOT EXISTS CORTEX_APPS.CHATBOT;

USE DATABASE CORTEX_APPS;
USE SCHEMA   CHATBOT;

-- Stage that holds the app source files
CREATE STAGE IF NOT EXISTS CHATBOT_STAGE
    DIRECTORY = (ENABLE = TRUE)
    ENCRYPTION = (TYPE = 'SNOWFLAKE_SSE');

-- Upload source files from your local machine:
--   PUT file://cortex_chatbot.py  @CHATBOT_STAGE overwrite=true;
--   PUT file://environment.yml    @CHATBOT_STAGE overwrite=true;

-- Create the Streamlit app
CREATE OR REPLACE STREAMLIT CORTEX_CHATBOT
    ROOT_LOCATION = '@CORTEX_APPS.CHATBOT.CHATBOT_STAGE'
    MAIN_FILE     = 'cortex_chatbot.py'
    QUERY_WAREHOUSE = 'COMPUTE_WH'
    COMMENT = 'Cortex-powered chatbot with multi-turn conversation history';

-- Grant access to other roles as needed
-- GRANT USAGE ON STREAMLIT CORTEX_APPS.CHATBOT.CORTEX_CHATBOT TO ROLE <role>;

-- Open the app URL
SHOW STREAMLITS;
