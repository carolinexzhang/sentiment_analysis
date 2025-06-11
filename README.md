# sentiment_analysis
Health-related sentiment analysis of popular Tweets using Excel VBA &amp; Python

This project analyzes approximately 5,000 health-related tweets using a combination of Excel VBA and Python for data cleaning, sentiment analysis, and descriptive analytics.

The core objectives of the project include:

1. Data Cleaning:
- Removal of duplicate tweets using Excel and a custom isDup VBA function.
- Implementation of logic to assess text similarity with a 70% threshold.

2. Sentiment Analysis:
- Python-based sentiment scoring system using a curated keyword list of positive and negative terms.
- Tweets are categorized as Positive, Negative, or Neutral using a custom VBA function.

3. Descriptive Analysis:
- Custom VBA functions mimicking Excel's COUNTIFS and AVERAGEIF to generate summaries by tweet group and topic.
- Identification of the group and topic with the highest sentiment score using a custom findMaxElement function.

4. Technologies Used:
- Excel (.xlsm) with VBA macros
- Python for sentiment scoring
- Textual analysis methods
- Data visualization and tabulation in Excel

Outputs:
- rawData – Original dataset with duplicates flagged
- processedData – Cleaned dataset with sentiment values and categories
- keywords – Positive and negative keyword lists
- analyzedData – Summary statistics including average sentiment and tweet counts by group and topic

Learning Outcomes:
- Advanced Excel/VBA scripting
- Integrating Python with Excel workflows
- Natural language processing basics
- Building custom analytics pipelines from raw textual data
