# swiss-scrabble

A swiss pairing system for Scrabble. based on https://github.com/gnomeby/swiss-system-chess-tournament/
That was the only python implementation of the swiss pairing system that I could find which adhered closely
to the specifications set by FIDE.

For scrabble unfortunately there isn't anything at all. The most popular scrabble pairing software, TSH does
not provide a complete implementation of the swiss pairing system. It's implementation is more like the monroe system
(King Of The Hill in scrabble speak)

At the moment the data is stored in spreadsheets using openpyxl and it can be exported into TSH and similarly TSH division files can be imported into our spreadsheets. Database support and a web interface will be added soon (via Django ORM)

