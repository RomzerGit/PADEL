# Padel ELO Ranking System

A system to track padel matches between a fixed group of players and calculate dynamic ELO rankings based on match results.

## Features

- **Excel-based data input**: Easy to add match results in an Excel spreadsheet
- **Dynamic ELO calculation**: Automatic rating updates for doubles matches
- **Player management**: Tracks unique players with starting ELO of 1000
- **Rankings dashboard**: Current ELO rankings, top 3 players, match history
- **Statistics**: Matches played, win rate, and more
- **Automation**: Python script recalculates ELO when new matches are added

## Installation

1. Install Python dependencies:
   ```bash
   pip install -r requirements.txt
   ```

## Usage

1. Run the script to create the initial Excel file:
   ```bash
   python padel_elo.py
   ```
   This creates `padel_matches.xlsx` with the required sheets.

2. Open `padel_matches.xlsx` and add match data to the "Matches" sheet:
   - Date: Match date (e.g., 2024-01-01)
   - Location: Optional location
   - Player 1, Player 2: Team 1 players
   - Player 3, Player 4: Team 2 players
   - Winning Team: 1 or 2

3. Run the script again to calculate ELO rankings:
   ```bash
   python padel_elo.py
   ```
   The script will update player ELOs and generate rankings in the "Rankings" sheet.

4. View results:
   - "Players" sheet: Individual player stats
   - "Rankings" sheet: Sorted ranking table
   - Console output shows top 3 players

## ELO System

- **Team Rating**: Average of both players' ELO ratings
- **K-Factor**: 32 (configurable in the script)
- **Rating Update**: Based on expected vs actual match outcome
- **Default ELO**: 1000 for new players

## Configuration

Edit the constants in `padel_elo.py`:
- `DEFAULT_ELO`: Starting ELO rating
- `K_FACTOR`: ELO sensitivity factor

## Future Enhancements

- ELO evolution graphs
- Head-to-head statistics
- Filtering by date/location
- Most frequent teammate/opponent tracking 
