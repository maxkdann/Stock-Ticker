# Stock-Ticker
A game I made in Excel using VBA and SQL replicating the classic board game Stock Ticker.

## Running the game
* Download the zip and extract the files 
* Open up the Excel file 

Using the Application
If you have never played the game, this is a manual explaining how the application works as well as the rules of the game. Click Play to begin the game.
 
Click Play to start the game
![image](https://user-images.githubusercontent.com/49764019/211650515-6ccbf6c0-cf8e-4572-89b3-ce6abb50fd2a.png)

Once you click Play you will be prompted to either create a new account by clicking “new User” or log into an existing save using “Existing User”.
 ![image](https://user-images.githubusercontent.com/49764019/211650536-4cef840f-1729-47d0-8c3e-8daa0dc56d07.png)

Click one of the two options
 
If creating a new account this form will pop up, as long as your username isn’t already in the database a new account will be created.
 ![image](https://user-images.githubusercontent.com/49764019/211650585-af684c4d-6db1-4a27-90b0-8b1cfe0d748d.png)

A new account will be created with the username and password you enter
If you’re logging into an existing account than you will be directed to login instead.
 ![image](https://user-images.githubusercontent.com/49764019/211650605-75f5cc91-7567-4036-a944-72be70c1cb7e.png)

If the username and password entered match the information stored in the database you be entered into your saved game.


All passwords are hashed using the SHA-1 hashing algorithm and then stored securely inside an access database.
 ![image](https://user-images.githubusercontent.com/49764019/211650620-32acc9d8-4d21-4463-b60a-32aa4ab2f4b8.png)

Passwords are stored securely in hashes instead of plain text.
Once you have created a new account or logged into an existing account (that still had time remaining) then you be directed into the game.
Game Rules
Start
A player starts the game with $5000, with this money they can choose to buy any of the following 6 stocks: Gold, Silver, Bonds, Oil, Industrial or Grain in increments of 500, 1000, or 2000. Each stock has an initial par value of 100 ($1 or 100 cents) and the player can decide to buy as many (or few) stocks as they would like. You will have 10 minutes to make as much money as possible and try to beat the high score. Remember: buy low, sell high.
 ![image](https://user-images.githubusercontent.com/49764019/211650645-7230b465-d639-4f18-aa00-321a76f3e1a8.png)

Each time the market opens this menu will pop up, the first two options take you to identical screens to buy and sell stocks, the third simply returns you to the game.
 ![image](https://user-images.githubusercontent.com/49764019/211650673-59d6f568-0d3b-4369-a54b-144805ce39e2.png)

To buy a stock, click the stock name and then the amount. If you have enough money the stock will be purchased. Clicking Close will return you to the game and clicking Return to Menu will bring you back to the previous screen where you can then elect to sell stocks as well.

The Play
The behaviour of the stocks is determined by the rolls of three dice. The first die specifies the stock, the second the amount (5, 10 or 20) and the third the action (up, down or dividend). The up and down movements are fairly obvious, with the specified stock moving in the specified direction the specified amount. 
 ![image](https://user-images.githubusercontent.com/49764019/211650691-20fd8f8f-93cc-4194-8b10-8ee405385dcb.png)

On the left, results of the dice roll are displayed. The roll button rolls all 3 dice at once.
 ![image](https://user-images.githubusercontent.com/49764019/211650711-2e7b1657-a778-4f76-96f0-b979960c5266.png)

The current values of stocks are displayed on this chart, which updates each time the dice are rolled.
If a dividend is rolled and the stock’s value is at or above par (above 100) then a dividend is awarded to any player owning that stock. The amount of the dividend is calculated as a percentage of the amount of stock the player owns. For example if the roll is Gold, 20, Dividend (with the value of Gold above par) and the player happens to own 1000 shares of Gold, then their dividend will be 20% (the rolled amount) of 1000 (the amount of stock the player owns) so the player would receive $200. What is important to note about dividends is the amount received does not change based on the current value of the stock therefore a player will receive the same dividend if the value of the stock is 100 or 190. 
  ![image](https://user-images.githubusercontent.com/49764019/211650736-2db141e4-7386-4392-85a4-12f52d8fd402.png)
![image](https://user-images.githubusercontent.com/49764019/211650751-15afa413-0884-484d-a32e-1cc3902eec04.png)

When a dividend is rolled and the stock is above par, your money increases.
 
If a stock reaches a value of 200, a stock split is declared. In a stock split, the value of the stock is reset to 100 and any amount the player owns of that specified stock is doubled. For example, if a player owns 500 Industrial and the value of Industrial reaches or exceeds 200, the value of Industrial will be reset to 100 and the player will now own 1000 shares of Industrial.
 ![image](https://user-images.githubusercontent.com/49764019/211650770-b43ea433-19b7-4086-bdb2-3b1f396b7974.png)

A stock that’s value goes above 199 gets its value reset to 100 and any stock you owned is doubled
If a stock reaches a value of 0, bankruptcy is declared. If a stock goes bankrupt, the value of the stock is reset to 100 and any amount the player own of that specified stock is reset to 0. For example, if a player owns 2000 Bonds and the value of Bonds reaches or falls below 0, the value of Bonds will be reset to 0 and the player will now own 0 shares of Bonds.
 ![image](https://user-images.githubusercontent.com/49764019/211650787-89802530-4e73-45d1-aadd-35bab1e2fb44.png)

A bankrupt stock gets its value reset to 100 but any stock you owned is reset to 0.
 
Every five turns the market opens. At this point (and only at this point), a player is allowed to buy or sell stocks. Any stock that has a value other than par will have an adjusted price based on its current value. For example, if the value of Gold has fallen to 90, 500 shares would now cost $450 (90% of 500). On the flip side, if Silver rises to 110, 500 shares would now cost $550 (110% of 500).
 ![image](https://user-images.githubusercontent.com/49764019/211650813-7b0b1b21-147f-41ae-8783-c3d6a3ce0bd4.png)

This table shows the prices of each stock in each increment available for purchase.
Finishing the game
At the end of the time limit, your final net worth is calculated based on the amount of cash you have plus the current market value of all stocks you own. A report is generated that compares your score with the high score.
