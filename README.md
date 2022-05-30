# colorful-conway
Using Excel for Conway's Game of Life

I’ve created Conway’s Game of Life in Excel, which is a mathematical game with only four rules:

1.	Any live cell with fewer than two live neighbours dies, as if by underpopulation.
2.	Any live cell with two or three live neighbours lives on to the next generation.
3.	Any live cell with more than three live neighbours dies, as if by overpopulation.
4.	Any dead cell with exactly three live neighbours becomes a live cell, as if by reproduction.

In the version called “User Choice”, you create the initial conditions, or “seed”, by filling in the cells you want with a color.  To do this quickly, select desired cells and press control+k.  To turn cells back to white, select desired cells and press control+r.  The game space is limited to the space within the rectangle border.  When you’re ready to play, just hit the button.  You can choose how many generations you want it to run.  Each generation lasts a little under a second, so choose wisely, or else you will have to use your task manager to end it if it lasts too long.  There are ways to set up initial conditions to create certain objects, like stable patterns, oscillating patterns, and “spaceships” that glide across the screen.  There’s a Wikipedia article about it:  https://en.wikipedia.org/wiki/Conway%27s_Game_of_Life.  Also, for fun, Google Conway’s Game of Life to see it happen in your browser.

In the one called “Random Beginning”, it will create a random seed when you hit the button.  You don’t create the seed in this one.  Just choose initial population density, as well as the number of generations.

These versions of Conway’s Game of Life change color, with colors reflecting which generation the cells come from.

Enjoy!
