makesummary
===========

A script I use to handle the finances of our team. Really not reusable for anyone else, but it is here for show nontheless.

Traverses a directory of team member files and produces an Excel XML-document.

Team member files have the following format:  
    Member name  
    Payment plan  
    Starting sum  
    date;label;sum         # transaction, zero or more  

Example:  
    John Tester  
    Junior player  
    1200  
    2012-06-11;Payment;600  
    2012-06-15;Prize money;1695  
    2012-06-28;Payment;1500  

Usage:

    $ make_summary_float.py directory-of-files

Will output to stdout so you probably want to redirect to some *.xml file.

Note that the different payment plans need to be entered in the source file where it is marked so.
