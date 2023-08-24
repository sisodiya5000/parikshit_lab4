class Match:
    def __init(self, location, team1, team2, timing):
        self.location = location
        self.team1 = team1
        self.team2 = team2
        self.timing = timing

class FlightTable:
    def __init(self):
        self.matches = []

    def add_match(self, match):
        self.matches.append(match)

    def search_by_team(self, team_name):
        matching_matches = [match for match in self.matches if team_name in [match.team1, match.team2]]
        return matching_matches

    def search_by_location(self, location):
        matching_matches = [match for match in self.matches if match.location == location]
        return matching_matches

    def search_by_timing(self, timing):
        matching_matches = [match for match in self.matches if match.timing == timing]
        return matching_matches

def main():
    flight_table = FlightTable()

    flight_table.add_match(Match("Mumbai", "India", "Sri Lanka", "DAY"))
    flight_table.add_match(Match("Delhi", "England", "Australia", "DAY-NIGHT"))
    flight_table.add_match(Match("Chennai", "India", "South Africa", "DAY"))
    flight_table.add_match(Match("Indore", "England", "Sri Lanka", "DAY-NIGHT"))
    flight_table.add_match(Match("Mohali", "Australia", "South Africa", "DAY-NIGHT"))
    flight_table.add_match(Match("Delhi", "India", "Australia", "DAY"))

    while True:
        print("Choose a search parameter:")
        print("1. List of all the matches of a Team")
        print("2. List of Matches on a Location")
        print("3. List of Matches based on timing")
        print("4. Exit")
        choice = int(input("Enter your choice: "))

        if choice == 1:
            team_name = input("Enter the team name: ")
            matches = flight_table.search_by_team(team_name)
            if matches:
                print("Matches for", team_name)
                for match in matches:
                    print(f"Location: {match.location}, Teams: {match.team1} vs {match.team2}, Timing: {match.timing}")
            else:
                print("No matches found for the specified team.")

        elif choice == 2:
            location = input("Enter the location: ")
            matches = flight_table.search_by_location(location)
            if matches:
                print("Matches at", location)
                for match in matches:
                    print(f"Teams: {match.team1} vs {match.team2}, Timing: {match.timing}")
            else:
                print("No matches found for the specified location.")

        elif choice == 3:
            timing = input("Enter the timing: ")
            matches = flight_table.search_by_timing(timing)
            if matches:
                print("Matches at", timing)
                for match in matches:
                    print(f"Location: {match.location}, Teams: {match.team1} vs {match.team2}")
            else:
                print("No matches found for the specified timing.")

        elif choice == 4:
            break

        else:
            print("Invalid choice. Please choose a valid option.")

if __name__ == "__main__":
    main()