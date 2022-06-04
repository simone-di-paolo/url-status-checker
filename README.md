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

<p text-align="center">
The script will create a new file with all the urls re-writed into the first column and into the second column of the file. Then in the third column you will have the history (if there is an history!) with a message like: <i>"Redirect status: 301"</i> (but it can be even a 302) got from the respone and into the fourth column you'll have a something like <i>"The redirect is correct - Status code: 200"</i> of some kind of errors.</p>

<h3>SETTING UP VARIABLES</h3>
  
  <p>Open the script with an editor like Notepad++ and edit the following variables as you prefer:
  <ul>
    <li><b>startFromWhichRow = 1</b>  # in case the first row of your Excel file has no rules but only column titles put 1, otherwise 0</li>
    <li><b>sheetNumber = 0</b>  # the index of the excel sheet you want to process, if doOnlyOneSheet is False it will do them all starting from this index</li>
    <li><b>doOnlyOneSheet = False</b>  # if for any reason you need to do one file sheet at a time put this to True (with capital T letter)</li>
  </ul>
  
  <p><b>N.B.</b> REMEMBER TO CHANGE ADD CORRECT PATHS TO THE FILES THAT YOU'LL READ AND WRITE!</p>
  
  <h3>INSTALLATION GUIDE</h3>
  
  <p>To make this script work you need to install Python and xlrd library.</p>
  <p>You can download the latest version of Python from those links for <a href="https://www.python.org/downloads/" target="_blank">Windows</a>, <a href="https://www.python.org/downloads/source/" target="_blank">Linux/UNIX</a>, <a href="https://www.python.org/downloads/macos/" target="_blank">MacOS</a>.</p>
  
  <p>If you are using .xlsx files, you'll need to install this specific version:</p>
  <pre>pip install xlrd==1.2.0</pre>
  <p>Otherwise, if you are using .xls files, then, you'll need to install the latest version:</p>
  <pre>pip install xlrd</pre>
  
  <p>You need to install even the xlwt library that will allow you to create and write into new .xlsx files and requests library too:</p>
  <pre>pip install xlwt
  pip install requests</pre>
  
  <p>Once done, navigate into your folder with cmd/terminal and launch (N.B. use your python version into the next command):</p>
  <pre>python url-status-checker.py</pre>
  
  <p>Now, you will find your new .xslx file into your folder destination (specified inside the script).</p>
</div>
