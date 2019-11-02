# No Database Files Present
This is where I stored my Microsoft Access `.mdb` files. I haven't included them for obvious data protection reasons. I describe the schema used here.

There were two database files; I decided at the time to split the user information out from the other data -- so there was `users.mdb` and `data.mdb`. Don't remember why.

## `data.mdb`
There were seven tables in this database, namely:

* `Attendance`, recording what members of staff were absent for a given day/period;
* `Cover`, recording what members of internal staff *covered* absent staff;
* `Departments`, listing the different departments the school had;
* `OCover`, listing staff who could be drafted in from the outside if needs be (an external agency?);
* `Periods`, listing the number of periods in each school day;
* `Rooms`, listing all the different rooms in the school; and
* `Timetables`, listing all internal teaching staff and their timetables.

The schema was horrendous, and I knew this when I started coding at the time. However, by the time I realised (I still remember that sinking feeling!), I couldn't go back and change everything and still meet the August 2006 deadline. Bummer. So I stuck with it. The design of the schema definitely impacted performance, and that is a valuable lesson I learned which I still carry with me to this day. You learn from your mistakes. This was an abomination. You will see why.

`Attendance`

| DAY | DATE       | USER | DEPT | 1_1  | 2_1  | 3_1  | ..._1 | 10_1 | 1_2 | 2_2 | ...  | 10_7 |
|-----|------------|------|------|------|------|------|-------|------|-----|-----|------|------|
| TUE | 12/09/2006 | 123  | 45   | NULL | NULL | NULL | ...   | NULL | A   | A   | ...  | NULL |

That's the schema with an example row. Ew! `2_1` represented 'Period 2, Day 1', with Day 1 being a Sunday, and Day 7 being the following Saturday. The system supported 10 periods in a day. When a member of staff was absent for a particular period, the corresponding entry would be set to `'A'`. Disgusting approach. And with this approach, you'd need to clear the database every week... Thankfully, I fixed that with PleaseTakes 2.

`Cover`

| FOR | COVERING | DAY | DAYDATE    | PERIOD | OCOVER |
|-----|----------|-----|------------|--------|--------|
| 123 | 456      | TUE | 12/09/2006 | 2      | NULL   |

This table records the members of staff who covered a class. Continuing the example from the `Attendance` table above, teacher `123` is absent period 2 on Tuesday, 12/09/2006. Teacher `456` covers their class. `OCOVER` would be set a value if external cover had been arranged; i.e. not from a teacher contracted to the school.

`Departments`

| ID | SHORT     | FULL        | DEPTID |
|----|-----------|-------------|--------|
| 10 | Maths     | Mathematics | 3      |
| 11 | Computing | Computing   | 17     |

A simple table representing all the different departments within the school. The idea here was to check to see if teachers in the absent teacher's department were free; these would be favoured for obvious reasons. Interesting how I didn't fully understand primary keys at the time; for some reason, I decided to introduce a further `DEPTID` field, when the `ID` field would have sufficed...

`OCover`

| LN  | FN   | TITLE | ENTITLEMENT | ENTIRE |
|-----|------|-------|-------------|--------|
| Doe | John | Mr    | 27          | 1      |

A table storing all available external cover staff. With `LN`, `FN` and `TITLE` self-explanatory, `ENTITLEMENT` represented the number of periods per week that that external cover member was entitled to cover per week, with `ENTIRE` indicating whether they were available the whole week in question. Remember, the database needed to be cleared every week; `Attendance` and `Cover` would be cleared. The SQL query that would calculate the number of periods taken up would subtract from the `ENTITLEMENT` value to give the number remaining. These staff were typically only called in when a teacher was off sick for a prolonged period of time (e.g. one or more weeks).

`Periods`

| TOTALS | DAYNAME | DAYNAMEFULL | DAYID |
|--------|---------|-------------|-------|
| 0      | Sun     | Sunday      | 1     |
| 7      | Mon     | Monday      | 2     |
| 7      | Tue     | Tuesday     | 3     |
| 6      | Wed     | Wednesday   | 4     |
| 7      | Thur    | Thursday    | 5     |
| 6      | Fri     | Friday      | 6     |
| 0      | Sat     | Saturday    | 7     |

Simple table, storing the number of periods in each day of the week. My school had seven on Mondays, Tuesdays and Thurdays, with six on Wednesdays and Fridays. This was a botch job; the session before (2005-2006), it had been six periods each day. In 2006, this was changed -- so I had to implement a hacky solution. Later I found that some schools used bi-weekly timetables! This solution would not have worked in those circumstances. Again, why did I need `DAYID`? Totally redundant given there was also a built-in `ID` field (not shown).

`Rooms`

| ID  | ROOMNO |
|-----|--------|
| 190 | D12    |

Simple. An `ID` for a room (didn't duplicate the field here!), and a human-identifiable room number. Interesting fact: `D12` was the room my computing teacher, Mr Phillips, used. Good times. Think it was renumbered `D13` not long after I left the school in 2008.

`Timetables`

| ID | LN  | FN   | TITLE | DEPT | CATEGORY | ENTITLEMENT | DEFROOM | 1_1  | 2_1  | ... | 10_1 | 1_2 | 2_2 | ... | 10_7 |
|----|-----|------|-------|------|----------|-------------|---------|------|------|-----|------|-----|-----|-----|------|
| 12 | Doe | Jane | Miss  | 10   | T        | 4           | F4      | NULL | NULL | ... | NULL | 2.3 | AH  | ... | NULL |

This was the big table! It contained the entire school's timetable, on a per-teacher basis. With hindsight, this was again a poor design choice -- I used the horrible `1_1`, `2_1`... approach again here. I should have split that off into another table (which I did in PleaseTakes 2 -- it's amazing what a university education can do for you!). `DEPT` matched with the `DEPTID` field in `Departments` (`10` in the example matched `Maths`). `CATEGORY` determined what grade the teacher was at -- `T` for just a teacher, `PT` for a Principal Teacher (of a given subject), `HT` for Head Teacher, etc. `ENTITLEMENT` determined how many classes a teacher would be allowed to cover each week; `DEFROOM` was their default, go-to room. Then followed the timetable information. For example, `1_2` indicates period `1` on day `2` (`MONDAY`), and they taught class `Maths 2.3`. Next period, they taught `Maths AH` (Advanced Higher). **I have nightmares about this table layout.** I don't really. But I know it was poor design.

## `users.mdb`
Users were split into either `Admin` or `User` users. `Admin` users could arrange cover and perform other tasks (like backing the system up), and `User` accounts were just for teachers to log in and see their cover. In the end, `User` accounts were not used; one of the requirements of the system was to be able to print paper slips which would be delivered in the morning to the unfortunate teachers who got a PleaseTake. The slip indicated what room to go to, who they were covering, and what subject it was. I still have the very first one that was printed off on Monday, September 11, 2006.

`Admin`

| ID | TTID | UN      | PW       | P1   | P2   | ACCLEVEL | TITLE | FN   | LN  | EMAIL          | DEPT | LASTLOGIN         |
|----|------|---------|----------|------|------|----------|-------|------|-----|----------------|------|-------------------|
| 2  | 12   | JaneDoe | password | NULL | NULL | 1        | Miss  | Jane | Doe | jdoe@email.com | 10   | 11/03/2008, 11:30 |

Lots of duplication here that an RMDBS would have been able to avoid, had I known better at the time. `TTID` matched to the `ID` in `Timetables`. But as I split the databases up, traditional relationships would not have worked! `Password` was unsalted, and stored in plaintext. Ouch. `P1` and `P2` was an additional security feature, providing the option to enter a six-digit PIN number to complement the password. Nobody used it. `ACCLEVEL` was either `1` or `2` -- `1` granting the ability to arrange cover, `2` allowing one to login and view cover for the school, but not arrange it. `DEPT` matched with `DEPTID` in `Departments`. Finally, `LASTLOGIN` was the date and time that the user successfully last logged in. A real time/date from my sample database was used in that example!

## `backup.mdb`
When the admin user cleared the database at the end of each week, the `Attendance` and `Cover` tables would be copied over to this database, prepended with the date range for that week. I have a nice example of a `backup.mdb` file with over three years' worth of cover information still in my archives.