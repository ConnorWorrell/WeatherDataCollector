import random
import time

winstreak = 0
rank = 25
stars = 0

wins = 0
losses = 0

for i in range(100):
    win = random.randint(1, 2)
    if(win == 1):
        wins = wins + 1
        stars = stars + 1
        winstreak = winstreak + 1
        if(winstreak > 2):
            stars = stars + 1
        if(rank <= 0):
            rank = 0
        else:
            rank = 25 - int(stars/5)

        print("Won Game,  Rank: " + str(rank) + " Stars: " + str(stars) + " Winrate: " + str(wins/(wins+losses)))
        time.sleep(.1)
    else:
        losses = losses + 1
        if (rank <= 0):
            rank = 0
        else:
            rank = 25 - int(stars / 5)
        if(rank <= 20):
            stars = stars - 1
        winstreak = 0
        print("Loss Game, Rank: " + str(rank) + " Stars: " + str(stars) + " Winrate: " + str(wins/(wins+losses)))
        time.sleep(.1)