# Blood on the Clocktower Almanac Creator

This project is used to create an easy to print almanac out of any script you desire in the same style as the Base 3 character almanacs. 
 
## Usage
There are 3 things this tool can do, and all 3 have command line interfaces so no code thought is necessary. The 3 things are

1) Add a character to the Almanac JSON: Copy the data over from the wiki so you can use a new character
2) Create a master almanac: Make an almanac with all characters loaded in the JSON file
3) Create a Script almanac: Admittedly the primary purpose. Input a name and a file location of your JSON script and it will provide an almanac with the name provided in the base directory. 

### Adding a character
When a new character is added, it needs to be added here. Luckily the interface is easy for that. Steps to do that
1) Go to the Wiki and download the image from there, and crop it so the edges are closer to the actual icon. Save this in the `Images/` directory under the name `Icon_<character_name_lowercase>.png`.
2) Start the program and include a response with a Y in it. (The program just looks for a 'y' or 'Y'. "Not yet" will add a character). 
3) Answer the prompts it provides, including the name, ability text, quote, etc. Do not include `"` characters in the surrounding the ability text or surrounding the quote. When it asks how many of something there are, answer with a number like `4` not `four`. 
A box is something a in the How to Run section inside a rectangle. If there is a horizontal line, there are multiple boxes. 
4) Create a master almanac and do a visual spot check of how much space it takes up. If more than one column on one page, cut down on information until it fits (examples would be first to go, followed by information that seems almost unrelated or spoon feeding spoons -- think Bounty Hunter entry). 

I will update this when new characters are released -- this mainly reminds me how to do this

### Creating a master almanac
On the first prompt, simply do not include a `y` or `Y` and it will fall through to ask if you want to make a master almanac. Enter in a prompt with a `Y` or `y` and it will create one. Give it ~10-15 seconds to operate (the process is a bit intense). 

The almanac will be saved in the top level directory under `Master Almanac.docx`

### Create a script almanac
On the first two prompts, simple do not include a `y` or `Y` and it will fall through to script almanac creation. 
1) First it will ask the name of the Script. Input it how you want it to appear on the cover. 
2) Enter the name of the file you want to base tha almanac off of, this file needs to be in the `Scripts/` directory.

At this time the script Almanac will be in the top directory of this project next to `Master Almanac.docx`

## Future Development

In the future I will work on some of the following things
1) Homebrew support. The ability to create almanacs in the TPI style without touching the main Almanac.json file
2) Different ways to slice up a master almanac (i.e. If you want just the Demons)
3) Whatever I come up with else! I'm probably the main user of this, so if I want something, I'll probably just write it in. Those first 2 I want. 

## Special notes
If something seems off about some of the inclusions, that's because there is. Some simply have no How to Run section, no examples, or a write up that's out of date of the ability change! I did not add any content to this project compared to the Wiki, only subtracted when needed. 