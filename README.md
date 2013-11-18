Datasheet.cfc
=============

Work with XLS files in Railo with POI

Background
----------

Since moving to Railo, I've yet to find a perfect solution to reading and writing XLS files.

After looking at POI in a bit more detail, I decided I would write my own component which makes use of Apache POI to read and write XLS files.

The plan is to start with the very basics, then to add additional functionality as needed.

Still to come
-------------

- Dates returned correctly
- asQueries() with good data
- asQueries() with messy data