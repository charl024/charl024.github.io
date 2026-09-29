This is the project repository for a Spin-a-Wheel topic chooser, written in HTML and Javascript (to be hosted on github.io).

What this application does is let the user upload (like, drag and drop) a spreadsheet file (let's say an excel file, but can also work with a csv) and turn each entry into an entry for this Spin-a-Wheel interactable. 

Users can spin this wheel. Upon a chosen wheel item, it is stored in a local session such that the user can see what items the wheel has chosen for them. The user has the choice to either remove that item from the wheel, or keep it in the wheel. 

The user is capable of changing several settings: how long (in seconds) for the wheel to spin before landing on an item, what seed to use (a string, that the user types into a text box), a "shuffle" option that shuffles the items in the wheel around.