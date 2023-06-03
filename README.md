<!-- PROJECT LOGO -->
<br />
<div align="left">
  <a href="https://github.com/simone-di-paolo">
    <img src="resources/img/sdp-logo-black.png" alt="Logo" width="80" height="80">
  </a>
</div>

<div align="left">
  <h3>SIMPLE REDIRECT TOOL GENERATOR FOR APACHE</h3>

<h3 dir="auto"><a id="user-content-what-are-vine-copulas" class="anchor" aria-hidden="true" href="#what-are-vine-copulas"><svg class="octicon octicon-link" viewBox="0 0 16 16" version="1.1" width="16" height="16" aria-hidden="true"></path></svg></a>What is this script for?</h3>

<p text-align="center">
    This is a simple script in Python that will check a series of urls from the first column of an .xlsx file, get the history from the response (redirects etc...) and will tell you, in a new .xlsx file, if the redirect is present (or not), if the page goes in 404 and if the redirected url is correct (based on the second column of the excel file that contains the list of the "final" urls).
</p>

<h3>SETTING UP VARIABLES</h3>
  
  <p>Open the script with an editor like Notepad++ and edit the following variables as you prefer:
  <ul>
    <li><b>Start from row</b>  # specify the first row of your .xlsx/.xls file that starts with a URL</li>
    <li><b>Sheet number</b>  # the index of the excel sheet you want to process, if "Selected sheet only" is unchecked it will do them all starting from this index</li>
    <li><b>Selected sheet only</b>  # if for any reason you need to do one file sheet at a time, flag this checkbox </li>
  </ul>
  
  <h3>INSTALLATION GUIDE</h3>
  
  <p>To make this script work you need to install Python, xlrd and xlwt libraries.</p>
  <p>You can download the latest version of Python from those links for <a href="https://www.python.org/downloads/" target="_blank">Windows</a>, <a href="https://www.python.org/downloads/source/" target="_blank">Linux/UNIX</a>, <a href="https://www.python.org/downloads/macos/" target="_blank">MacOS</a>.</p>
  
  <p>Move into project folder with git bash or cmd and launch this command to easy install all the required libraries:</p>
  <pre>pip install -r requirements.txt --user</pre>
  
  <p>Just in case you are using .xls (and not .xlsx) files, then, you'll need to install the latest version manually:</p>
  <pre>pip install --upgrade xlrd</pre>

  <p>Once done, make sure you're in the correct folder with cmd/terminal and launch (N.B. use your python version into the next command):</p>
  <pre>python gui.py</pre>
  
  <p>Now, setup all needed inside the brand new GUI and start the script.</p>
  <p>After that, you will find your new results.xslx file into your folder destination (specified inside the GUI).</p>
  
  <p>NB: an .exe of the script is available but your OS will probably block it. The best way to run the script is buy command line as described above but if you want, you can try to whitelist the .exe and run it in the easiest way possible.</p>
</div>

## UPDATES
<h3>03-06-2023 - <a href="https://github.com/simone-di-paolo/url-status-checker/releases/tag/v0.1.3">Version 0.1.3</a></h3>

Features:
- Added the requirements.txt file in order to easily install all the required packages
- Added the .exe of the script to the release

<h3>02-06-2023 - <a href="https://github.com/simone-di-paolo/url-status-checker/releases/tag/v0.1">Version 0.1</a></h3>

Features:
- Added a first version of a GUI

<img style="width: 300px" src="https://github.com/simone-di-paolo/url-status-checker/assets/24905857/8a52ed6e-8f2a-4c61-aa5f-9faba92ae019"/>
