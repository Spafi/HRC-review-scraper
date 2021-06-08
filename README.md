<br />

  <h3 align="center">Review Scraper & Mail Sender</h3>
    <br />
  <p align="center">
    This app is a web scraper and mail sender that's currently used in a restaurant. It scrapes the reviews page of a restaurant on a delivery platform and sends the management team emails when they receive a bad rating to make recovery with their customers as fast as possible.
    <br />
    <br />
    <br />
  </p>

<!-- TABLE OF CONTENTS -->
<details open="open">
  <summary>Table of Contents</summary>
  <ol>
    <li>
      <a href="#about-the-project">About The Project</a>
      <ul>
        <li><a href="#built-with">Built With</a></li>
      </ul>
    </li>
    <li>
      <a href="#getting-started">Getting Started</a>
      <ul>
        <li><a href="#prerequisites">Prerequisites</a></li>
        <li><a href="#installation">Installation</a></li>
      </ul>
    </li>
  
  </ol>
</details>

<!-- ABOUT THE PROJECT -->

## About The Project

![Product Name Screen Shot][product-screenshot]
In the image above are shown the three possible log messages.
<br>
<br>

This application was created for the management team of a restaurant. It scrapes the reviews page from a delivery website for the restaurant and sends an email to the team when they receive a review rated under the selected rating value.

The main benefits of the application:

- It sends an email with all needed details as soon as the restaurant receives a bad review, so the management team can act fast in making recovery with their customers.
- Easy to integrate and use, from utilizing already existing communication methods in the company (Email)

<br>

![Product Name Screen Shot 2][product-screenshot-2]
The screenshot above is the email model that is received and we have the following:

- Order time and date
- Rating
- Customer's name
- Customer's phone number as a link that redirects to the caller app of the phone or laptop
- Order number as a link to the review page on the external website
- Review message (if any)
- A template for external mails to receive a voucher for the customer
- A link to an empty new email, with predefined fields for receiving before mentioned vouchers.

<br>

Being used with another website, many things can go wrong. The application is set to automatically shut down after five consecutive unsuccessful connection requests, but not before alerting the users trough an email.

### Built With

- [Python](https://www.python.org/)
- [BeautifulSoup](https://www.crummy.com/software/BeautifulSoup/)
- [Tkinter](https://docs.python.org/3/library/tkinter.html)

<!-- GETTING STARTED -->

## Getting Started

To get a local copy up and running follow these simple steps.
<br>
<b>NOTE! Being made for a specific company you will not be able to run it locally as is because you need the credentials of the company, but you can browse the code. </b>

### Prerequisites

To be able to run the project locally you'll need Python 3.

### Installation

1. Install project requirements
   ```sh
   pip install -r requirements.txt
   ```
2. Run the project

### Exporting as .exe file

1. Install with pyinstaller
   ```sh
    pyinstaller --onefile --noconsole review-scraper.py
   ```
2. Run .exe file located in `distro/`

<!-- MARKDOWN IMAGES -->
<!-- https://www.markdownguide.org/basic-syntax/#reference-style-links -->

[product-screenshot]: images/screenshot.png
[product-screenshot-2]: images/screenshot-2.png
