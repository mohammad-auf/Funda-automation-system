# Funda Bot Scraper Alert

The project was developed based on a client requirement to build an automated monitoring system for property listings on Funda.nl.

The client maintained a growing Google Sheet containing property addresses with the following fields:

* Street
* House Number
* Addition
* Postal Code
* City

The goal of the system was to automatically check daily whether any of these addresses appeared on Funda as properties for sale and send instant notifications when matches were found.

## Sample Output
Here is an example of the notifications sent by the Telegram bot when matches are found, along with the daily summary:

![Telegram Bot Output](screenshot.png)

To solve this, a complete Python-based automation system was developed with the following functionality:

### Core Features Implemented

* Automated scraping of all active property listings from Funda.nl
* Full pagination handling for continuous data extraction
* Extraction of detailed property information including:
  * Street
  * House number
  * Addition
  * Postal code
  * City
  * Listing URL

### Smart Matching System

An efficient local matching engine was built to avoid performing thousands of individual search requests.

The matching logic uses:

1. Postal code + house number (+ addition if available)
2. Street + house number + city as a fallback

This approach allows the system to process thousands of addresses efficiently while minimizing requests and reducing the risk of IP blocking.

### Dynamic Google Sheets Integration

The scraper automatically reads address data directly from Google Sheets, allowing the client to continuously add new addresses without changing the code or restarting the application.

The system was designed to support scalability from:

* 2100+ addresses initially
* Up to 2500+ addresses and beyond

### Notification System

Integrated Telegram notifications for:

* Newly matched properties
* Scraping/parsing errors
* System execution alerts

### Logging & Monitoring

Implemented a complete logging system to track:

* Start and end time
* Total listings processed
* Total matches found
* Errors and exceptions

### Performance & Stability

* Request throttling added to prevent rate limiting and IP blocking
* Optimized scraping workflow for VPS deployment
* Duplicate prevention system to avoid repeated alerts for the same listing

### Automation & Deployment

* Fully automated daily execution using cron jobs
* VPS-ready deployment for platforms such as Hostinger
* Clean, scalable, and maintainable Python codebase

The final system successfully delivers an efficient and fully automated property monitoring solution capable of handling large datasets with reliable daily notifications and minimal manual intervention.
