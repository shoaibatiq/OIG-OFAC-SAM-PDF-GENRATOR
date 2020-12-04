
# OIG, OFAC, SAM Report PDF Generator
 [![forthebadge made-with-python](http://ForTheBadge.com/images/badges/made-with-python.svg)](https://www.python.org/)  [![selenium ](https://img.icons8.com/color/48/000000/selenium-test-automation.png)](https://selenium-python.readthedocs.io/)<br> 
 <img src="https://i.imgur.com/466F7P8.png" alt="xlrd" height="100px" width="100px">
 ![Open Source Love svg1](https://badges.frapsoft.com/os/v1/open-source.svg?v=103)
  ![Ask Me Anything !](https://img.shields.io/badge/Ask%20me-anything-1abc9c.svg)
[![GitHub license](https://img.shields.io/github/license/Naereen/StrapDown.js.svg)](https://github.com/Naereen/StrapDown.js/blob/master/LICENSE)

The Bot reads data from excel file and serach that data on 
- https://exclusions.oig.hhs.gov/
- https://sanctionssearch.ofac.treas.gov/
- https://www.sam.gov/

and genrates PDFs of the result.


### Built With
* [Python](https://www.python.org/)
* [xlrd](https://www.crummy.com/software/BeautifulSoup/bs4/doc/)
* [pywinauto]()





### Prerequisites

### Installation
1. Clone the repo
```sh
git clone https://github.com/Zeeshanahmad4/Facebook-Automation-bot-with-Multilogin-and-Proxies.git
```

2. Install python packages
```sh
pip install selenium
pip install xlrd
pip install pywinauto

```

<!-- USAGE EXAMPLES -->
## Usage

You can edit path of data file PDF download location of each site in settings.py.
You need to run indiviual module for each site.

