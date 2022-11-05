# text_extract_tools
A GUI to paste text in, which can extract certain patterns and put them in a template excel file.

But very specific to certain texts at the moment.
I wish to make it more general with options to choose which patterns are enabled and what the output will be.
And also to save these as presets.

How it works now:
  Run extract.py
  Paste text in the textbox (at the moment hardcoded: emails, specific usernames and specificly formatted full names.)
  Press extract.
  
These steps load a template excel file.
Extract data from text box.
Write these to a new excel file based on the template.
