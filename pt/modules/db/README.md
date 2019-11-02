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
| TUE | 12/11/2006 | 123  | 45   | NULL | NULL | NULL | NULL  | NULL | A   | A   | NULL | NULL |

That's the schema with an example row. Ew! `2_1` represented 'Period 2, Day 1', with Day 1 being a Sunday, and Day 7 being the following Saturday. The system supported 10 periods in a day. When a member of staff was absent for a particular period, the corresponding entry would be set to `'A'`. Disgusting approach. Thankfully, I fixed that with PleaseTakes 2.