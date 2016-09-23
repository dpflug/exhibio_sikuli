from pprint import pprint
import csv
import re
alphabetic = re.compile('\D+')

#Let's speed things up.
Settings.MoveMouseDelay=0

def save_schedule():
    """Opens, saves schedules as csv, closes Excel.
       
       Assumes you have Outlook open to the email with the Excel attachment.
       Actually makes a lot of assumptions, but that's life with Sikuli.
    """
    # It's a fuzzy match, so this should work:
    doubleClick(Screen(1).wait("1454000341475.png", 60))
    # Also, the amount of "fuzz" is configurable. 
    wait("1454000411454.png", 15)
    type("s", Key.CTRL)
    wait("1454000434509.png", 3)
    type("e")
    wait("1454000492008.png", 3)
    type("\n")
    wait("1454000509864.png", 5)
    type("cschedule.csv\tc\n")
    if exists("1454000531481.png"):
        type("y")
    wait("1454000626904.png", 3)
    type("y")
    type(Key.F4, Key.ALT)
    wait("1454000580160.png", 3)
    type("n")


def init_exhibio():
    """Let's pull up the Sikuli webpage."""
    switchApp(r"C:\Program Files\Google\Chrome\Application\chrome.exe")
    wait("1454000837040.png", 30)
    # Maximize
    type(Key.SPACE, Key.ALT)
    wait(0.5)
    type("x")
    wait(1)
    type("L", Key.CTRL)
    type("exhibio.circuit5.org\n")
    if exists("1454000857364.png", 5):
        type("\texhibio\texhibio\n")


def findDay(today, tries):
    """Gets us into today's edit page."""
    titles = {"MONDAY": "1454001256863.png",
              "TUESDAY": "1454001265351.png",
              "WEDNESDAY": "1454001276416.png",
              "THURSDAY": "1454001284800.png",
              "FRIDAY": "1454001299576.png"}
    def aux(tries):
        """This function isn't strictly necessary.

           It tries to find the day's title in the list,
           and scrolls through when it fails.

           However, since we're using the shorter list mode,
           all the days fit in one page. However, I'm leaving
           it here for reference.
        """
        try:
            click(titles[today])
        except FindFailed:
            if tries > 0:
                click("1439576487948.png")
                aux(tries - 1)
            else:
                raise

    click(wait("1454001130719.png", 30))
    click(wait("1454001155262.png", 5))
    click("1454001165554.png")
    waitVanish("1454001211570.png", 5)
    aux(tries)
    click(wait("1454001323454.png",3))

def pick_slide(n):
    """Pull up slide n."""
    dropdown = wait("1454002257685.png", 2).offset(60, 4)
    click(dropdown)
    # Each page is 13px below the previous. The first page is 15px from the drop-down box.
    click(dropdown.offset(0,2 + 19 * n))
    # There isn't a visual indicator for when the slide is loaded, AFAICT. I'm just hoping 2 seconds is enough. :)
    wait(2)


# Ok. Let's get our data.
def parse_judge_schedule():
    """This massages our schedule data out of the csv created by save_schedule().

       It assumes we only have 2 categories of people: Judges and non-judges.
       A lot of the gymnastics are in handling the freaky formatting we end up with
       due to the original Excel file being decorated/intended for humans.

       row[0] is the judge's name
       row[4] is the time
       row[6] is the event or note
       row[8], if it exists, is the room/courtroom
       row[10] is the floor
    """
    cs = csv.reader(open(r'C:\Users\ablane\My Documents\cschedule.csv'), dialect='excel')
    today = cs.next()[8]
    cs.next()
    schedules = {'judges': [], 'other': []}
    current_judge = ''
    out_of_order = False
    current_position = 'judges'
    for row in cs:
        if alphabetic.match(row[0]):
            # When names are out of order, they're either a visiting judge or some other official.
            if current_judge and (out_of_order or ord(row[0][0]) < ord(current_judge[0])):
                out_of_order = True
                if popAsk("Is {} a judge?".format(row[0])):
                    current_position = 'judges'
                    sort_judges = True
                else:
                    current_position = "other"
            current_judge = row[0].title()
            if current_judge == "Williams":
                # Judge Ritterhoff-Williams wants to appear as such
                current_judge = "RitterhoffWilliams"
        # Don't include mediations or internal things
        if row[10] and not row[6] == 'Mediations' and row[6].find("Personnel") == -1:
            if row[8].strip() == "Jury Assembly Room":
                # JAR doesn't need a floor #?
                schedules[current_position].append((current_judge, 
                                                    row[4], 
                                                    row[6].strip().replace(" cont'd", ''), 
                                                    row[8].replace("bly Room",'bly')))
            elif row[8]:
                # Standard courtroom, 2D - 2nd floor for example
                schedules[current_position].append((current_judge, 
                                                    row[4], 
                                                    row[6].strip().replace(" cont'd", '').replace(' Docket', ''), 
                                                    row[8] + " - " + row[10] + " Floor"))
            elif row[10] and row[6]:
                # Check with security
                schedules[current_position].append((current_judge,
                                                    row[4], 
                                                    row[6].partition("-")[0].strip().replace(" cont'd", '').replace(' Docket', ''), 
                                                    row[10] + " Floor Security"))
            elif row[6]:
                # These are just notes. The space is to allow my tab to the next field to work in exhibio.
                schedules[current_position].append((current_judge, ' ', row[6].strip(), ' '))
        
    # Sort list of judges
    schedules["judges"].sort(key=lambda k: k[0])
    # and others
    schedules["other"].sort(key=lambda k: k[0])
    
    return(today, schedules)

def fill_table(data):
    rows = len(data) - 1
    type(Key.HOME)
    for row_num, row in enumerate(data):
        columns = len(row) - 1
        for column_num, column in enumerate(row):
            type(Key.DELETE * 5, Key.CTRL)
            type(Key.END, Key.SHIFT)
            print(row_num, column_num, rows, columns)
            if row_num < rows or column_num < columns:
                type("\t")

def enable_slide():
    """Makes sure current slide is active."""
    click(wait("1454001871349.png", 5))
    
    click(wait("1454001900776.png", 5))
     
    # We can be pretty exact about this.
    Settings.MinSimilarity = 0.85
    disabled = wait("1454001914606-1.png", 2).offset(225,0).exists("1454003185224.png")
    if disabled:
        click(disabled)
        click("1454003297692.png")
    #click("1440002819045.png")
    click("1454003308591.png")
    type(Key.PAGE_DOWN)
    type(Key.UP)
    # Wait on their little animation
    wait(1)
    # Return to default similarity
    Settings.MinSimilarity = 0.7

def fill_schedules(sched_data, judges=True):
    """Fills our schedule pages with our schedule data.

        TODO: Fix (split?) this function to work if there's multiple pages of non-judge schedule.

       Arguments:
       sched_data: List of pages. Pages are lists of schedule data. Schedule data is lists of columns in the table.
       judges: True if we're filling in the judges' schedule
    """
    data = list(paginate_schedule(sched_data, 9))

    for page_num, page in enumerate(data):
        if judges:
            # The schedules start on the 3rd slide, and enumerate starts at 0
            pick_slide(page_num + 3)
        else:
            # Everyone else only goes on the last slide.
            pick_slide(find_last_slide())
        enable_slide()
        click("1454005449444.png")
        type(Key.HOME, Key.CTRL)
        click(wait("1454005306463.png",3))
        type(Key.DOWN * 2)
        fill_table(page)
        if not popAsk("Look good?", "Page Finished"):
            popup("I apologize. I've obviously gotten confused. Tell David: #7874, dpflug@circuit5.org.")
            try:
                click("1454001685093.png")
            except:
                popup("Something has gone wrong and you'll have to clean up after me, too. :( Sorry.")
            return
        save = click("1454001697396.png")
        wait("1454001711696.png", 30)
        type(Key.ENTER)
        type(Key.HOME, Key.CTRL)

def find_last_slide():
    """Tries to OCR final slide info, falls back to asking the human."""       
    page_list = find("1454001744398.png").offset(42, 4).grow(110, 2).text()
    try:
        return int(re.search(r' of ?(\d+)\)', page_list).group(1))
    except AttributeError:
        return int(input("How many slides are there today? (What's the highest number in the \"(Page 1 of ...)\" part?)"))
    

def paginate_schedule(sched, per_page):
    for start in xrange(0, len(sched), per_page):
        yield sched[start:start+per_page]

save_schedule()
today, schedules = parse_judge_schedule()
init_exhibio()
findDay(today, 1)
fill_schedules(schedules["judges"])
if schedules["other"]:
    fill_schedules(schedules["other"], False)
