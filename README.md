Datasheet.cfc
=============

Work with XLS files in Railo with POI

Background
----------

Since moving to Railo, I've yet to find a perfect solution to reading and writing XLS files.

After looking at POI in a bit more detail, I decided I would write my own component which makes use of Apache POI to read and write XLS files.

The plan is to start with the very basics, then to add additional functionality as needed.

Environment
-----------

Built and tested on Railo 4.1.2.003, POI 3.8.

Coding style
------------

As an experiment, Datasheet.cfc is a script based component with ~~no use of semi-colons which makes use of~~ Railo specific features.

~~Why? To see how it feels to write script with no ; So far it feels great!~~

There are too many bugs in Railo when leaving semi-colons off the ends of statements so for now they'll be used.

Still to come
-------------

- Dates returned correctly
- asQueries() with good data
- asQueries() with messy data