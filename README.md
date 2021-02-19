# AstroPixels
VBA Excel code for gathering the ephemeris from Astropixels.com service

The Excel-Macro tool gathers and sanitizes the Solar Syatem ephemeris computations from the Astropixels.com service managed by Fred Espenak, who is an author of these numerical data. Because all these calculations are HTML-based and places losely between each other, you could just paste them in Excel as a whole string. From now, you just need select all computational data and copy them (Ctrl+C) to the clipboard. The "AstroPixels" tool will cater for the rest.
By using the "AstroPixels" macro, you can save plenty of time, assuming, that you are going to utilize the Astropixels.com ephemeris for some further calculations. By having all of them sorted in MS Excel it's quite easy to progress further with any formulas.

The major features of the "Astropixels" tool are:
- Sanitizing ephemeris on the Solar System objects (Sun, Moon and planets),  
- Calculating decimal declination for the Solar System objects (Sun, Moon and planets) ,
- Calculating distance to the Sun [km],
- Indicating perigeum and apogeum monents throughout the pasted year (by highlighting with the background colour),
- Cleaning the calendar of celestial events, including highlighting the most important ones throughout the year, included in the legend key next to,
- Cleaning the Great Annual Lunar Standstills 2001-2100 with marking the max and min lunar declination across this period
- Tidying the "Closest Moon Perigees (Super Moons) 2001-2100 table

The "Astropixels" tool doesn't cover:
- Earth's perihelion and aphelion times,
- Solstices and equinoxes,
- Sky event almanacs in time zones other than GMT (Greenwich Mean Time),
- Phases of the Moon,
- Closest Full Moon perigees,
- Passages of the Moon through nodes
- Monthly lunar standstills
- Full Moon at perigee (Super Moon) table, except of the closest perigees

These feaures can be subject for further tool development. However their lack in this tool means, that the relevant computations have been produced already in the "Ephemeris" section of the Astropixels.com service and visitors can find them quickly.

The tool comes in version 1 (v1.0). 

Getting the data is a simple process, which has been shown in the "Readme" sheet. You must copy the relevant sections to the clipboard only. No need to paste it to the Excel. Just click "GET DATA" and the tool will do the rest for you. 

The macro doesn't advance of saving option. If you are happy with the data, you must save or export it on your own.

The code is subject of copyright. Some elements come from stackoverflow.com and extendedoffice.com. I am open for any suggestions of changes, alterations and modifications, as far as they lead to the tool developent.


							
							
							
							
							
							
							
