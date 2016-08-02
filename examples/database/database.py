"""
Copyright (C) 2014-2016, Zoomer Analytics LLC.
All rights reserved.

License: BSD 3-clause (see LICENSE.txt for details)
"""
import sqlite3
import os
import xlwings as xw


def playlist():
    """
    Get the playlist content based on the ID from the Dropdown
    """
    # Make a connection to the calling Excel file
    wb = xw.Book.caller()
    sht = wb.sheets.active

    # Place the database next to the Excel file
    db_file = os.path.join(os.path.dirname(wb.fullname), 'chinook.sqlite')

    # Database connection and creation of cursor
    con = sqlite3.connect(db_file, detect_types=sqlite3.PARSE_DECLTYPES | sqlite3.PARSE_COLNAMES)
    cursor = con.cursor()

    # Get PlaylistId from ComboBox
    playlist_id = wb.api.ActiveSheet.OLEObjects("ComboBox1").Object.Value

    # Database query
    cursor.execute(
        """
        SELECT
        t.Name AS Track, alb.Title AS Album,  art.Name AS Artist, t.Composer
        FROM PlaylistTrack pt
        INNER JOIN Track t ON pt.TrackId = t.TrackId
        INNER JOIN Album alb ON t.AlbumId = alb.AlbumId
        INNER JOIN Artist art ON alb.ArtistId = art.ArtistId
        WHERE PlaylistId = ?
        """, (playlist_id,))

    # Get the result and column names
    col_names = [col[0] for col in cursor.description]
    rows = cursor.fetchall()

    # Clear the sheet and write the column names and result to Excel
    sht.range('A9').expand().clear_contents()
    sht.range('A9').value = col_names
    if len(rows):
        sht.range('A10').value = rows
    else:
        sht.range('A10').value = 'Empty Playlist!'

    # Close cursor and connection
    cursor.close()
    con.close()


def combobox():
    """
    This populates the ComboBox with the values from the database
    """

    # Make a connection to the calling Excel file
    wb = xw.Book.caller()
    source = wb.sheets['Source']

    # Place the database next to the Excel file
    db_file = os.path.join(os.path.dirname(wb.fullname), 'chinook.sqlite')

    # Database connection and creation of cursor
    con = sqlite3.connect(db_file, detect_types=sqlite3.PARSE_DECLTYPES | sqlite3.PARSE_COLNAMES)
    cursor = con.cursor()

    # Database Query
    cursor.execute("SELECT PlaylistId, Name FROM Playlist")

    # Write IDs and Names to hidden sheet
    source.range('A1').expand().clear_contents()
    source.range('A1').value = cursor.fetchall()

    # Format and fill the ComboBox to show Names (Text) and give back IDs (Values)
    # TODO: implement natively in xlwings
    combo = "ComboBox1"
    wb.api.ActiveSheet.OLEObjects(combo).Object.ListFillRange = \
        'Source!{}'.format(str(source.range('A1').expand().address))
    wb.api.ActiveSheet.OLEObjects(combo).Object.BoundColumn = 1
    wb.api.ActiveSheet.OLEObjects(combo).Object.ColumnCount = 2
    wb.api.ActiveSheet.OLEObjects(combo).Object.ColumnWidths = 0

    # Close cursor and connection
    cursor.close()
    con.close()

