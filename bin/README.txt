First, the boring stuff.
DISCLAIMER: This software (PT2SQLite) is provided "AS IS" and any expressed or implied warranties, including, but not limited to, the implied warranties of merchantability and fitness for a particular purpose are disclaimed. In no event shall the regents or contributors be liable for any direct, indirect, incidental, special, exemplary, or consequential damages to unicorns (including, but not limited to, procurement of substitute goods or services; loss of use, data, or profits; or business interruption). I love pie.
However caused and on any theory of liability, whether in contract, strict liability, or tort (including negligence or otherwise) arising in any way out of the use of this software, even if advised of the possibility of a big boo-boo.

Then, the README.

THIRD-PARTY FILES
1. COMCTL32.ocx: this file holds the ancient secrets of progress bars. I could do without but it's so much shinier with these progress bars !
2. sqlite3.dll: this is a wrapper for the SQLite engine. Basically, it's used to created the SQLite database. It's based on the official SQLite source files

REGISTRY KEY
1. One is used to store the DataServer path for next time. That's it.

HOW TO USE
(0. Make sure you got space on your drive!)
1. Launch the program (a double click will suffice)
2. Browse and select your dataserver folder
  2.1 The program WILL POUT if you select a folder that does not contain the userdata and warehouse folders!
3. Press the start button
4. Watch the neat progress bar fill up. The first one is for the userdata folder, the second one is for the warehouse folder.
  4.1 NOTE: the progress bar aren't incremented per file but per folder (the 0 to 255 folder)
  4.2 Pressing the Stop button will stop the scan (you didn't expect that did you ?)
5. Once the scan is finished, collect your precioooous file. The file extension is .ptsqlite and it's created in the same folder as the program. The name is the date and time of when you _started_ the scan, following this format: yyyy-mm-dd-hh-nn-ss
  - yyyy : year, 4 digits
  - mm : month, 2 digits
  - dd : day, 2 digits
  - hh : hours, 2 digits
  - nn : minutes, 2 digits (yes nn! mm was taken already ): !)
  - ss : seconds, 2 digits
6. Open your database file using a SQLite browser and start querying. If you need some documentation about SQLite queries, there's a website for that, please do RTFM.

SAFETY
1. The tools only READS within the DataServer folder. Open, read, close. It should not cause any harm to your files. As a matter of fact, it's impossible unless while you run it, the planets align and a giant beaver comes from nowhere and rapes your hard drive.
2. The only time the

THIRD-PARTY SOFTWARE
1. I recommend SQLite Database Browser. Free, open-source, simple.
  - http://sqlitebrowser.sourceforge.net/index.html
  - download is here: http://sourceforge.net/projects/sqlitebrowser/
  
UNKNOWN BEHAVIOR
1. I didn't run it on a big DataServer folder (10k+ .dat files). I don't know how it will react. The program should be good as it reads the files one by one. However I don't know if it'll be slowed down if the SQLite database grows to gigantic proportions.
2. I have no clue what size the SQLite database would be on a big Dataserver folder. For 1417 items, the file was 362kB. That's ~260B per record. So 1GB would hold something like 4 billion items. Whoopee doo.

BUGS
1. There's no default value for the DataServer path at first launch when there's no registry key
2. It fairs poorly if you're running the database file on a network share. Keep it local or you might miss a few records
3. The stop detection could be a bit more accurate. Currently it can only stop once it's done scanning a folder (not userdata or warehouse, the 0 to 255 folders)