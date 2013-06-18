makesummary
===========

A script I use to handle the finances of our team. Really not reusable for anyone else, but it is here for show nontheless.

Traverses a directory of team member files and produces an Excel XML-document.

Team member files have the following format:
Member name
Payment plan
starting sum
date;label;sum         # transaction, zero or more
2012-06-11;Payment;500 # example transaction

Usage:

    $ make_summary_float.py directory-of-files

Will output to stdout so you probably want to redirect to some *.xml file.
