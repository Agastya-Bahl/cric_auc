import requests
import json
import csv
import re
import time
import httpx
import gspread
from google.oauth2.service_account import Credentials
from gspread_formatting import *
from google.colab import drive

# Define border format
border_format = Borders(
    top=Border("SOLID"),
    bottom=Border("SOLID"),
    left=Border("SOLID"),
    right=Border("SOLID")
)

def get_column_letter(n):
    result = ""
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        result = chr(65 + remainder) + result
    return result

def get_data(event_id, score_dict, team_choice):
    url = f"https://www.sofascore.com/api/v1/event/{event_id}/innings?nocache={int(time.time())}"

    with httpx.Client(http2=True) as client:
        response = client.get(url)
        data = response.json()

    if "innings" not in data:
        return None
    
    catch_dict = {}
    innings = data["innings"]
    for inning in innings:
        if inning["battingTeam"]["shortName"] == team_choice:
            choice = "batting"
        elif inning["bowlingTeam"]["shortName"] == team_choice:
            choice = "bowling"
        else:
            choice = None
        compute_innings(inning, score_dict, catch_dict, choice)

    for k,v in catch_dict.items():
        score_dict[k] = score_dict.get(k, 4) + v * 8 + (4 if v >= 3 else 0)
    return data

def compute_innings(inning, score_dict, catch_dict, choice):
    bat_team = inning["battingTeam"]["shortName"]
    bowl_team = inning["bowlingTeam"]["shortName"]
    if choice != "batting":
        for bowler in inning["bowlingLine"]:
            if bowler["playerName"] not in score_dict:
                score_dict[bowler["playerName"]] = 4
            compute_bowler(bowler, score_dict)
    for batsman in inning["battingLine"]:
        if batsman["playerName"] not in score_dict and choice != "bowling":
            score_dict[batsman["playerName"]] = 4
        compute_batsman(batsman, score_dict, catch_dict, choice)
    
    return (bat_team, bowl_team)

def compute_bowler(bowler, score_dict):
    name = bowler["playerName"]
    overs = bowler["over"]
    economy = (bowler["run"] / overs) if overs > 0 else 8
    wickets = bowler["wicket"]
    maidens = bowler["maiden"]
    score = wickets * 25 + maidens * 12 + economy_score(economy, overs) + wicket_bonus(wickets)
    score_dict[name] = score_dict.get(name, 4) + score

def economy_score(economy, overs):
    if overs >= 2:
        if economy < 5:
            return 6
        elif economy < 6:
            return 4
        elif economy < 7:
            return 2
        elif economy < 10:
            return 0
        elif economy < 11:
            return -2
        elif economy < 12:
            return -4
        else:
            return -6
    else:
        return 0

def wicket_bonus(wickets):
    if wickets == 3:
        return 4
    elif wickets == 4:
        return 8
    elif wickets >= 5:
        return 16
    else:
        return 0

def compute_batsman(batsman, score_dict, catch_dict, choice):
    name = batsman["playerName"]
    runs = batsman["score"]
    fours = batsman["s4"]
    sixes = batsman["s6"]
    balls = batsman["balls"]
    sr = 100 if balls == 0 else ((runs * 100) / balls)
    score = runs + fours * 1 + sixes * 2 + sr_bonus(sr, batsman["player"], balls) + duck_check(runs, batsman["player"], batsman["wicketTypeName"] != "Not out", balls) + run_bonus(runs)
    if choice != "bowling":
        score_dict[name] = score_dict.get(name, 4) + score
    wicket_type = batsman["wicketTypeName"]
    if wicket_type != "Not out" and choice != "batting":
        compute_wicket(wicket_type, batsman, score_dict, catch_dict)

def sr_bonus(sr, player, balls):
    if player["position"] != "B" and balls >= 10:
        if sr < 50:
            return -6
        elif sr < 60:
            return -4
        elif sr < 70:
            return -2
        elif sr < 130:
            return 0
        elif sr < 150:
            return 2
        elif sr < 170:
            return 4
        else:
            return 6
    else:
        return 0

def duck_check(runs, player, is_out, balls):
    if player["position"] != "B" and runs == 0 and is_out and balls > 0:
        return -2
    else:
        return 0

def run_bonus(runs):
    if runs < 30:
        return 0
    elif runs < 50:
        return 4
    elif runs < 100:
        return 8
    else:
        return 16

def compute_wicket(type, batsman, score_dict, catch_dict):
    if type == "Bowled" or type == "LBW":
        score_dict[batsman["wicketBowlerName"]] += 8
    elif type == "Caught":
        if batsman["wicketCatchName"] not in catch_dict:
            catch_dict[batsman["wicketCatchName"]] = 1
        else:
            catch_dict[batsman["wicketCatchName"]] += 1
    elif type == "Stumped":
        score_dict[batsman["wicketCatchName"]] = score_dict.get(batsman["wicketCatchName"], 4) + 12
    elif type == "Run out":
        score_dict[batsman["wicketCatchName"]] = score_dict.get(batsman["wicketCatchName"], 6) + 6

def get_participant_points(score_dict, gw_no, best_xi_dict, missing_set, num_players=11, folder = ""):
    participant_dict = {}

    with open(f"{folder}/teams/gw{gw_no}teams.csv", mode='r') as file:
        reader = csv.reader(file)
        key = "feewd XI"
        role = None
        player_lst, point_lst, role_lst = [], [], []

        for line in reader:
            for text in line:
                text = text.strip()  # Remove spaces
                text = re.sub(r'\s+', ' ', text)  # Normalize spaces

                if text.startswith('*'):  # New team detected
                    if player_lst:  
                        participant_dict[key] = list(zip(player_lst, point_lst, role_lst))  
                    key = text[1:].strip()
                    player_lst, point_lst, role_lst = [], [], []
                elif text.lower() in ["batsmen", "all-rounders", "bowlers"]:  
                    role = text.lower()
                elif text:  # This is a player
                    player_name = text[:-5] if text.endswith("(WK)") else text
                    if player_name in score_dict:
                        player_lst.append(text)
                        point_lst.append(score_dict[player_name])
                        role_lst.append(role)  
                    else:
                        missing_set.add(player_name)  # Missing player warning

        if player_lst:  
            participant_dict[key] = list(zip(player_lst, point_lst, role_lst))

    # Sorting teams by player points (descending order)
    for team, players in participant_dict.items():
        # Categorizing players based on roles
        batsmen, bowlers, all_rounders, wks = [], [], [], []

        for player, points, role in players:
            if "(WK)" in player:
                wks.append((player[:-5], points))
            elif role == "batsmen":
                batsmen.append((player, points))
            elif role == "bowlers":
                bowlers.append((player, points))
            elif role == "all-rounders":
                all_rounders.append((player, points))

        # Sorting all categories by points
        batsmen.sort(key=lambda x: -x[1])
        bowlers.sort(key=lambda x: -x[1])
        all_rounders.sort(key=lambda x: -x[1])
        wks.sort(key=lambda x: -x[1])

        # Selecting the best XI
        conf_bat, conf_bowl, conf_ar, conf_wk = [], [], [], []
        tc = 0
        batsmen_count, bowlers_count = 0, 0

        # Ensure at least one WK is in the team
        if wks:
            conf_wk.append(wks.pop(0))  # Pick best WK
            tc += 1

        wk, wk_points = conf_wk[0] if conf_wk else ("N/A", 0)
        wk_capt, wk_vc = False, False

        # Merge all categories for best selection
        all_players = batsmen + bowlers + all_rounders + wks
        all_players.sort(key=lambda x: -x[1])  # Sort by points
        for player, points in all_players:
            player_to_add = player
            if tc == 1 and wk_points < points:
                player_to_add = player + " (C)"
            elif tc == 1:
                player_to_add = player + " (VC)"
                wk_capt = True
            elif tc == 2 and wk_points < points:
                player_to_add = player + " (VC)"
            elif tc == 2:
                wk_vc = True
            if tc == 11:
                break  # Stop when we have 11 players

            # Add batsmen (including WKs)
            if ((player, points) in batsmen or (player, points) in wks) and batsmen_count < 5:
                conf_bat.append((player_to_add, points))
                batsmen_count += 1
                tc += 1

            # Add bowlers
            elif (player, points) in bowlers and bowlers_count < 5:
                conf_bowl.append((player_to_add, points))
                bowlers_count += 1
                tc += 1

            # Allow all-rounders to fill gaps
            elif (player, points) in all_rounders:  
                conf_ar.append((player_to_add, points))
                tc += 1

        if wk:
            if wk_capt:
                wk = wk + " (C) (WK)"
            elif wk_vc:
                wk = wk + " (VC) (WK)"
            else:
                wk = wk + " (WK)"
                
        best_xi = conf_bat + conf_ar[:5-batsmen_count] + [(wk, wk_points)] + conf_ar[5-batsmen_count:] + conf_bowl
                
        best_xi_dict[team] = best_xi
    return best_xi_dict

def output_participant_points(best_xi_dict, missing_set, game, update_sheet, folder = ""):
    max_per_row = 4
    if len(missing_set) > 0:
        print("MISSING PLAYERS:")
        for player in missing_set:
            print(player)
    # Writing output to CSV and calculating standings
    standings = {}
    with open(f"{folder}/points/game{game}points.csv", mode='w', newline='') as file:
        writer = csv.writer(file)

        teams = list(best_xi_dict.items())
        team_chunks = [teams[i:i + max_per_row] for i in range(0, len(teams), max_per_row)]  # Split into groups of 4

        for chunk in team_chunks:
            # Write team names and points in a single row
            row = []
            for team, best_xi in chunk:
                team_points = sum(int(p[1]) for p in best_xi)
                row.extend([team, team_points, ""])  
                standings[team] = team_points
            writer.writerow(row)

            # Get max number of players in the chunk
            max_players = 11

            # Write players row by row
            for i in range(max_players):
                row = []
                for _, best_xi in chunk:
                    if i < len(best_xi):
                        row.extend([best_xi[i][0], best_xi[i][1], ""])  # Player name, points
                    else:
                        row.extend(["N/A", 0, ""])  # Empty cells for alignment
                writer.writerow(row)

            writer.writerow([])  # Blank row for spacing

    standings = dict(sorted(standings.items(), key=lambda item: item[1], reverse=True))

    print("\nSTANDINGS:")
    for rank, (team, points) in enumerate(standings.items(), 1):
        print(f"{rank}) {team}: {points}")
    return standings

def print_to_sheets(game, data, standings, folder = ""):
    # Authenticate with Google Sheets API
    creds = Credentials.from_service_account_file(f"{folder}/credentials.json", scopes=["https://www.googleapis.com/auth/spreadsheets"])
    client = gspread.authorize(creds)
    SHEET_ID = "1AEn2LG9bfTAQdZonbeNe5xg6EI9LXfpf5gcVJ5yD0eM"
    doc = client.open_by_key(SHEET_ID)
    no_sheets = len(doc.worksheets())
    if no_sheets < game:
        sheet = doc.add_worksheet(title=f"GAME {game} TABLE", rows="1000", cols="26") 
    else:
        sheet = doc.get_worksheet(game-1)
    
    # Use the first sheet
    rankings = []
    if game == 1:
        rankings = sheet.get_all_values()[:8]
    else:
        rankings = doc.get_worksheet(game-2).get_all_values()[:8]
    columns = list(map(list, zip(*rankings)))
    columns[1][1:], columns[game + 1][1:] = standings.keys(), standings.values()
    rankings = list(map(list, zip(*columns)))
    rankings[0][game + 1] = f"Game {game}"
    rankings[0][game + 2] = "TOTAL"
    for i in range(7):
        rankings[i+1][game + 2] = sum(int(rankings[i+1][j]) for j in range(2, game + 2))
    sheet.update(values = rankings, range_name = "A1")
    sheet.update(values = data, range_name = "B15")
    # Apply to a range (e.g., A1:D10)
    format_cell_range(sheet, f"A1:{get_column_letter(game + 3)}8", CellFormat(borders=border_format))
    format_cell_range(sheet, f"A1:{get_column_letter(game + 3)}1", CellFormat(textFormat=TextFormat(bold=True), horizontalAlignment='CENTER'))


def main(game, update_sheet = True, folder = ""):
    to_open = f"{folder}/ids/game{game}ids.csv"
    with open(to_open, mode='r') as file:
        data = csv.reader(file)
        event_ids = {}
        for lines in data:
            event_ids[lines[0]] = {"gw_no": lines[1].strip(), "team_choice": lines[2].strip() if len(lines) > 2 else "B"}

    player_pos_dict = {}
    team_dict = {}
    score_dict = {}
    best_xi_dict = {}
    missing_set = set()

    for event_id, e_dict in event_ids.items():
        data = get_data(event_id, score_dict, e_dict["team_choice"])
        # print(data)
        score_dict = dict(sorted(score_dict.items(), key=lambda item: item[1], reverse=True))
        missing_set.clear()
        get_participant_points(score_dict, e_dict["gw_no"], best_xi_dict, missing_set, folder = folder)

    standings = output_participant_points(best_xi_dict, missing_set, game, update_sheet, folder = folder)

    
    with open(f"{folder}/calcSheet.csv", mode = 'w') as file:
        writer = csv.writer(file)
        for k, v in score_dict.items():
            writer.writerow([f"{k}: {v}"])

    if update_sheet:
        with open(f"{folder}/points/game{game}points.csv", newline='') as f:
            data = list(csv.reader(f))
        print_to_sheets(game, data, standings, folder = folder)
            

if __name__ == "__main__":
    drive.mount('/content/drive')
    main(1, update_sheet = True, folder = "drive/MyDrive/Cric Auc")
