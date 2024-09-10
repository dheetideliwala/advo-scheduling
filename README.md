# ASU Devils' Advocates Scheduling AppsScript
**Purpose:** This project includes two scripts with different input types from Welcome2College.com. I created and used both versions during the 2022-2023 academic year to schedule over 100 tour guides for over 100 tours monthly. 

## ShiftScheduling.js
**Input:** guide selected shifts
![Sample Input from w2c](/README-images/Input.png)

### Setup
![Roster tab](/README-images/Dashboard.png)
The `Roster` tab displays tour guide names, their status in the organization, and the number of tours scheduled weekly and for the month.
Tour guides should only be scheduled for maximum 2 tours in the month and maximum 1 tour in a week.
Newbies and LOA members will not be scheduled for tours, and their status is referenced from the `Newbies` and `LOA` tabs.
- After newbies earn their polo certification, they will be scheduled as active tour guides (2 tours/month) once their date of certification is entered in the `Newbies` tab.

The `New Month` button resets the dashboard for the month, clearing all tour count numbers

### Scheduling Process
1. Copy the input into the tab corresponding to the correct week.
2. In the cell to the right of each tour date for the week, enter the number of tour guides needed.
3. Click the `Schedule Guides` button. This will select guides for all the tours in the week. Selected tour guides will be highlighted in green and have an X placed next to their name.
![Scheduled Guides Sample](/README-images/Tour1.png)
![Scheduled Guides Sample](/README-images/Tour2.png)
4. The Dashboard will automatically be updated with tour counts.
![Tour Counts](/README-images/TourCount.png)



## TimeAvailability.js & TimeScheduling.js
**Input:** available time slots
![Sample Input from w2c](/README-images/Input2.png)

### Setup
The `Roster` tab displays tour guide names, their status in the organization, whether they have entered their availability in w2c, and the number of tours scheduled for the month.

The `Availabilities` tab marks a tour guide as "Yes" or "No" for every 15 minutes from 8am to 6pm Monday-Friday.

![Weekly Availability](/README-images/Weekly%20Availability.png)
The `Advo Availability` tab displays the number of tour guides available every 15 minutes Monday-Friday (data from the `Availabilities` tab). This is helpful for the Advancement Chair and tour advisor when scheduling advancements and additional tours.

### Scheduling Process
1. Copy the input into the `W2C Availability` tab.
2. Click the `Delete Old Availability` button to change all cells in `Availabilities` to "No"-- essentially resetting everyone's availabilities.
3.  Click the `Add New Availability` button. This will input "Yes" in `Availabilities` whenever tour guides are free to give tours.
![Availability Entered](/README-images/Availabilities.png)
4. In the `Tour Schedule` tab, create the monthly tour template from the original Tour_Creater.xlsx. Adjust names as needed.
5. The Dashboard will automatically be updated with with tour counts based on the number of times a tour guide's name is found in the `Tour Schedule` tab.