	  CompApp Timed ShutDown Clock


Thank you for using CompApp Timed ShutDown Clock.
This is one of many programs by CompApp and I hope 
you find this program usefull and error free. If 
there are any problems, please e-mail me at the 
address below with your problem, and I'll try to
fix it for you. 
- Eric

License
For freeware use only. You cannot legally sell this 
program or make any changes to this file if you
intend to distribute this product. Usual copyright
laws apply. (see below after version history)


Created By  : Eric O'Sullivan
Company     : CompApp Technologies
Contact     : DiskJunky@hotmail.com
Web Address : http://www.compapp.co-ltd.com

========================================================
========================================================
========================================================




Updates - Version History
========================================================
v 6.4.293
Code overhauwl done. About screen code redone to be 
faster, intelligent saving, more protection against
ini file data corruption (can be annoying). The program
in general has been stripped of unnecessary code and
has more internal checks to make sure it is as stable
as possible.


v 6.3.287
The mouse cursor is now animated to move to the "Yes" 
button when the computer asks to shut down the computer.
New Option: Prevent other applications from closing 
windows. This means that you can now stop others from 
shutting down windows until the clock shuts it down 
itself. You can override this by using the end-task window 
or turning the option Off.
A small bug with customized title bar sizes when 
switching between analogue and digital modes is now 
fixed.
A major overhaul on code effeciency and graphical info.
Unnecessary code has been stripped down and deleted
from the program for faster loading times and better
program execution.
The program now uses less cpu time. The main cause for
this was checking for the idle time. As stated above,
this was updated for more efficient use of resources.

v 6.2.264
New option: In the Shutdown Options screen, there is a
new set of options called Idle Time. With these options
the program can now shutdown your computer after it has 
been idle for a specified amount of time (like a 
screensaver).
New option: Start Minimized. This will make the program
start as a minimized icon instead of a window. 
New option: Change System Time. This will let you change
the computers' local time. Note: on Win NT you will need
the appropiate security rights to change the time.
New option: Snap Window. This will snap the clock to the
side of the desktop like winamp.
About screen has been redone (no flicker). Fixed bug
when loading invalid background picture.
The clock now records where you left it before the program
terminated. This means that you can move it to a part of 
the screen and the clock will always start there until you
move it.
I finally managed to rid the Time box of its flicker. 
I know I've said this before, but there was always
a little tell-tale flicker. Anyway, even this has (to my 
relief) been totally eradicated.
Also some fine tuning to the code has been done to speed
up the program and make it more aesthetically pleasing.

v 6.1.187
Shut down related bug fixes. About screen has been re-
done. New option: Always On Top. This will put the clock
on top of what ever screen is currently active.

v 6.0.178
I've added two new screens to the program. A Password
screen and a Colour Schemes screen. Also, you can now
double-click on the system-tray icon to show the clock
instead of using the menu (the option is still there
though). Some animation for when you try to exit the
program when the Timed Shut-down is still active, helps
let you know that the program hasn't stoped running.
Also a password decryption bug was fixed (slight error
during coding). 
A new option was added aswell. The 'Run At Startup'
option lets you customize whether or not the program
will activate itself every time windows is started.
That annoying flicker is also fixed (to my relief).

- Eric

========================================================
v 5.9.163
Program now finds the licensed user in about 2 seconds,
instead of 10 minutes. New password protection. Double
click on sys-tray icon to show clock.

v 5.8.153
Some bug fixes. Default background setting changed. File
path validation in code for background.

v 5.7.151
Increase .ini ledgability and a new option on the 
"Shut Down Options" screen. You can choose your shut down
method (this increases compatability). Also faster code
in search for licensed owner (by approx 60 times).

v 5.6.143
Major bugs in the setup program fixed and new "Advanced" 
option added.

v 5.4.134
A 'Second' counter in shut down window and overflow error
fixed.

v 5.3.128
The program now find the licencsed owner of the machine.

v 5.2.127
New option ; "Re-load SysTray Icon" comes in handy if
for some reason the system tray is re-set.

v 5.1.115
A few minor bugs from the new background screen fixed.

v 5.0.108
Again, I've improved the data entry on the shut-down
options screen. Also I've no included a 'Background'
option on the main menu. You can set a picture or logo
into the background of the clock. There is a preview 
screen to help you decide wich picture you want and 
you can turn off the background to your clock background 
colours (the colours menu).
The jump in version numbers (from .70 to .108) was mainly
caused by me tracking down an error that only occured in
the .exe, hence I had to make a lot of .exe's to find and
cure the problem.

- Eric
========================================================
v 4.0.70

You can now change the colours of the various parts of
clock. The colour options are on the main menu and are
self-explanitory. I've also added an icon the the startbar
to help close and show the clock in case of emergency.

- Eric
========================================================
v 3.4.61

You can now set the shut down times for any particular
day. Also in version 3.4.61, the time entry system was
improved so most of what's described for V2.3.45 
Shut down options doesn't really apply anymore.
{ReadMe help updated} 

Please don't screw around with the SetClock.ini file or 
the program could do some weird things.

- Eric

========================================================

CompApp Timed ShutDown Clock v2.3.45
Created by Eric O'Sullivan

========================================================
CompApp owns the copyright of this program but takes
no responsability for any muck-ups to any part of
your computer (including, but not limited to,
not saving any data in applications active before
shutting down the computer, improper operation
as a result of someone screwing around with the
.ini file etc.) as a result of running this program 
(usual copyright stuff here....).
========================================================

Anyway, to use the CompApp clock, just run it by
clicking on the shortcut in your Programs folder
on the Start menu.

INSTALLING
Versions greater then v2.0.34
Just run the Setup.exe after unzipping the file.
If you using WinZip, it will let you install
without unzipping.

v 2.0.34
For those who are new to this, this means that
you click on the 'Start' button, go to 'Find'
and click on 'Files or Folders...'. You should
see a flashing line in a white box with 'Named'
written beside it. Type in 'CompApp*.exe' and 
press the [RETURN] key. If you see a file called
'CompAppClock.exe' in the white box that just 
appeared, then doule-click on it to run the clock.

v 6.1.187
Basically the same procedure as described in 
v 2.0.34, but search for a file called 
"Clock Setup*.exe" and run the program to install the
clock.

========================================================

========================================================


KNOWN PROBLEMS
There really is only one known problem. The clock has
been known to cause slowdown with other vb (visual 
basic) programs. This is due to the clock using off-
screen bitmaps. It keeps a background constantly which
slows the computer down by taking up ram (not a whole
lot, say 100-200 Kilobytes, plus program size).


USING
To use the clock, just start it and right-click 
anywhere on the progrma that just appeared. You
should see a menu appear with;

(i) 	'Analogue'
(ii)	'24 Hour'
(iii)   'Idle Shut Down'
(iv)	'Timed Shut Down'
(v)	'Shut Down Options'
(vi)	'Snap Window'
(vii)	'Background On/Off'
(viii)	'Background Options'
(ix)	'Colour Schemes'
(x)	'Password Options >'
(xi)	'Advanced Options >'
(xii)	'About'
(xiii)	'Exit'


(i) Analogue
If this option has a tick beside it, then you
should be seeing a clock face on top of a digital
clock. Selecting this option again turns it off.


(ii) 24 Hour
If this option has a tick beside it, then it means
that the digital clock is running in 24 hour mode.
Selecting this option again will turn it off and
the digital clock will display 12 hour mode.


(iii) Idle Shut Down
If this option is checked, then the clock will shut
down the computer after a certain amount of time idle.
You can set the amount of time after which the computer
will shut down the computer by selecting 'Shut Down 
Options' from the menu. See section (v) for more.


(iv) Timed Shut Down
If this option is not ticked (default), then it
means that the compter will not shut down at the
per-specified time (see Shut down options for
default shut down time). Selecting the option 
when it's not ticked will activate the timed
shut down.

IMPORTANT : The program will still be active, if
not visible, when you try to close the prorgam
and the 'Timed Shut Down' option is ticked. To
close the program untick the 'Timed Shut Down'
option BEFORE you close the program. You can 
also close the program totally, if you right
click (use the right mouse button to click) on
the clock icon in the startbar (right hand side),
and select "Quit". Much simpler.


(v) Shut Down Options...
When you click on this option you should be
presented with another screen. 

Timed ShutDown At
To set the time of the shut down, please enter the
time you want to shut down at IN 24 HOUR MODE only.
This is important. If you enter a value greater
than the valid value (eg entering '50 in the 
'Hour' box), then the value is automaticaly taken
as zero. Zero in the hour box is 12 O' clock 
midnight.

Default ShutDown Delay
This option needs a bit of explaining. This means
that when 'Confirm Shutdown' is turned on, the 
Confirm Shutdown screen is activated (at the time 
you set for the computer to shutdown, it'll ask for 
Yes or No), the computer will automaticaly shut down 
after the specified number of seconds. This gives you
the oppertunity to cancel the shut down of the 
computer, but also, if no answer is given within
the specified time (in seconds), the computer will
shut down straight away WITHOUT WARNING.
If you enter a value greater than 59 in the 'Seconds'
box, it will take the value within 60. eg, if you 
enter 73, then it wil take is as '60'. Be careful with 
this. If however, you unchecked the 'Shutdown On/Off' 
box, the computer will never shutdown.

Shut Down Method
This is mainly explaned in the window itself, but I'll
repeat it hear. This option was added because in some
computers the hardware did not support some types of
shutdown options. Example : You select the Shut Down
method, but instead of shutting down completly, you
get a screen saying "It is now safe to turn off your
computer". To get araound this you can select the
"Power Down" option.
Note :  if the Shut Down option works properly, then,
by default, some of the other options will not work on
your computer. I'd would advise you to experiment to see
which one works best.

Idle Time
This option, if turned on, will shutdown your computer
after your computer has been idle for a specified amount
of time.
You can turn this option on by making sure there is a 
check mark in the "Idle Shutdown On" box.
To set the time, simply select the appropiate values from
the Hours and Minutes combo boxes (drop-down boxes). The
maximum time allowed is 23:59 (23 hours, 59 minutes). The
minimum is 1 minute. If, for some reason, the time was less
than one minute, it is automatically changed to one hour.


(vi) Snap Window
When there is a check mark next to this option, it means
that the clock will "stick" to the sides of the screen or
taskbar (like winamp).


(vii) Background On/Off
This turns the background picture on or off. The 
picture can be set in (vi) Background Options.


(viii) Background Options
This option will show a screen with a review 
screen. To set a picture, just click on the
'Browse...' button and select a picture. You
should now see a preview of what the picture
will look like. You can stretch or tile the
picture ot fit the clock or you can have no 
effects for the picture.
To put the picture into the clock, just click
on the 'Set Picture' button.


(ix) Colour Schemes
Just select the colour you want to change and
click. You will see a window with a list of 
colours and a preview pane in it. Select the 
colour you want and click '...'. Another
window will appear. Select the colout you
want and click 'Ok'.The colour is changed.
To make the colour permenant, click on 
'Set Scheme". To save the colours as a scheme,
click "Save Scheme". When you want to view a
previous scheme, just select it from the menu.


(x) Password Options
  1) 'Enter/Change Password'
     Pretty much what it says. Enter or change the 
     password. If no password is currently set, 
     then, the program will ask you for one.

  2) 'Password Enabled'
     Only available when either the option is NOT
     checked or after you have entered the password.

  3) 'Lock Menu'
     Only available when the password is active and
     the password has been entered. This option will
     de-select some of the menu options until the 
     password is entered again.


(xi) Advanced Options
  a) Always On Top
  b) Prevent Shut Down
  c) Run At StartUp
  d) Start Minimized
  e) change System Time
  f) Re-Load System Tray Icon  
  g) Shut Down Computer
  h) Re-Start Computer
  i) Log-Off User
  j) Power Down

  a) Always On Top
	If checked, then the program will force itself on
	top of whatever windows is currently active. Eg,
	suppose the option is checked and you just opened
	Microsoft Word. The program will still be visible,
	even though your using Word.

  b) Prevent Shut Down
	This option, if checked, will prevent other 
	applications from closing windows. This means that
	you can use the clock to keep windows active until
	you want it to be shut down. However, for saftey
	reasons, you can bypass this with the end-task
	dialogue box.

  c) Run At StartUp
	If checked, then the program will run every time
	windows starts. If unchecked, the program will
	NOT run every time windows starts and will require
	you to run it from the start menu.

  d) Start Minimized
	If this option is checked, the program will appear
	as a icon on top of the Start-bar, even if you 
	restart the program.

  e) Change System Time
	This option will bring up a window with three boxes.
	The first contains the hour you want to change to, the
	second contains the minute you want to, and the third
	box contains the second you want to change to.
	To set the new time, just click on the "Set" button

  f) Re-Load System Tray Icon
	Just click on this option if you can't see the 
	clock icon on the right hand side of the start
	bar.

  g) Shut Down Computer
	Just what it says - it shuts the computer down.

  h) Re-Start Computer
	Same again, it only does what it says - restarts 
	the computer.

  i) Log-Off User
	This will log you off the current user (even if
	the computer is not set up for it). 
	NOTE : windows is still active, but Explorer.exe
	       is not loaded so you can't see icons and 
	       the Start bar etc.

  j) Power Down
        This option (only on some computers), will power
        down the computer (turn off, totally) after windows
        shuts down.


(xii) About
Show information about the clock and where to
contact me if you have any trouble. Click 'Ok'
to get rid of the window.


(xiii) Exit
Guess what this does :) 

Note: the program will not really exit if the 
"Timed Shutdown" option is checked, but you
can close the program properly by right-clicking
on the clock icon in the system tray (the box on the
righthand side of the startbar), and selecting "Quit"


========================================================

========================================================

Well that's it really. I hope you enjoyed this 
clock and please look out for other CompApp 
applications, created by Eric O'Sullivan.

For more information, please contact:
	DiskJunky@hotmail.com
or visit :
	http://www.compapp.co-ltd.com


