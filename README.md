# PleaseTakes
Back when I was 15/16 years old, [my high school](https://blogs.glowscotland.org.uk/er/MearnsCastle/) computing teacher asked me to write some software with the aim of helping the Senior Management Team (SMT) of the school. This was in June 2006; by August 2006 (after the summer break), I came back with this.

It's called **PleaseTakes**, a mashup of the words *please* and *take*, i.e. *please take this class*. If a teacher could not take a class because of illness or some other reason, the SMT needed to manually lookup a copy of the printed timetable for every other member of staff in the school, see who was free for the required time, whether they were in the school or not, and tell them to 'babysit' the class without a teacher. This system changed all that; it fully automated the process. It was a godsend, and I got a lot of success out of it.

The system was used until August 2010, when I replaced it with PleaseTakes 2. You'll find a lot of issues and stupid things in this code. Remember, I was 15/16 years old when I wrote this. I didn't know what object-orientated programming was; nor did I really have an understanding for writing code with the smallest time complexity. There are some horrendous loops in this code which, with hundreds of teachers, took two or three minutes to complete! At least I engineered in a nice loading box for that delay...

It's ancient. It's useless. But I wanted to put up an archived repository to show the world that *I can code*, and by looking at later repositories, you'll see how my ability to code evolved and improved.

## Lack of Databases
I haven't included any database files in this repository for obvious reasons. I wrote a README file in the `pt/modules/db` directory that provides the basic schemas that I used. They were horrible and ill-thought out. But then again, I had no idea how to do better at the time. It's a miracle this worked at all.

## Technologies Used
This was developed using ancient technologies, too. I hadn't quite grasped object-orientated programming at this stage (that came in 2009 at university, when I was taught Java and object-orientated concepts by [Rob Irving](http://www.dcs.gla.ac.uk/~rwi/) -- probably the course I have learnt the most from during my time on this earth), so I resorted to old-school VBScript.

* Microsoft Internet Information Services (IIS) server, version 6.0
* Microsoft Active Server Pages (Active Server Pages), version 3.0
* Microsoft JET Database Engine, version 4.0

Yes, this was a Windows-based project; I didn't know much about the open-source world at the time, having grown up with MS-DOS and Windows. And yes, the database was a Microsoft Access database. I had no idea about SQL Server at the time. And yes, everything is mashed into different `.asp` files. Oh god, the humanity! At least I tried to split things up a bit by separating commonly-used stuff into include files (`.inc`). There's that, at least.

**This is very much my first major outing into the software development world. I didn't have formal training for anything at the time; all of this was entirely self-taught, and pre-computing-science-at-university. Be mindful of that when you look at the code!** 🤓