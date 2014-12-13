/*Summary
Script which determines the best possible combination of abilities for groups of each size
Best combinations are determined by the average success rate of the group after doing every mission
*/

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel;
using System.IO;
using System.Data;

namespace OptimalFollowerGroups
{
	//Class which represents an individual follower and their abilities
	public class Follower
	{
		public string name;
		public int[] abilities = new int[2];

		public string[] traits = new string[3];
		public string race;

		public double missions_completed = 0;
		public double missions_completed_2 = 0;
	}

	//Represents a mission which contains ability counters as well as trait counters
	public class Mission
	{
		public int[] abilities;
		public string trait;
		public Follower[] best_group;
		public double prescence = 0;

		public Mission Clone()
		{
			return new Mission() { abilities = (int[]) abilities.Clone(), trait = trait, best_group = best_group };
		}
	}

	//Class which represents a group of followers
	//Can be any size and calculates results for missions
	public class Group
	{
		public List<Follower> followers = new List<Follower>();
		public List<double> mission_results = new List<double>();
		public double total_ability_used { get { return mission_results.Sum(); } }

		//All possible combinations of followers
		List<Follower[]> _combinations_followers = null;
		public List<Follower[]> combinations_followers
		{
			get
			{
				if (_combinations_followers == null)
				{
					SetupCombos();
				}
				return _combinations_followers;
			}
			set { _combinations_followers = value; }
		}

		public void SetupCombos()
		{
			_combinations_followers = new List<Follower[]>();
			for (int i = 0; i < followers.Count; i++)
				for (int j = i; j < followers.Count; j++)
					for (int k = j; k < followers.Count; k++)
						_combinations_followers.Add(new Follower[] { followers[i], followers[j], followers[k] });
		}

		//Find best result for mission
		public void RunMission(Mission input_mission)
		{
			mission_results.Add(0);
			input_mission.best_group = null;
			//Go through each combination of followers in this group
			foreach (Follower[] follower_combo in combinations_followers)
			{
				var mission = (Mission) input_mission.Clone();
				double prescence = 3;
				foreach (Follower f in follower_combo)
				{
					if (mission.abilities[f.abilities[0]] != 0)
					{
						f.missions_completed++;
						mission.abilities[f.abilities[0]]--;
						prescence += 3;
					}
					if (f.abilities[1] > -1 && mission.abilities[f.abilities[1]] != 0)
					{
						f.missions_completed++;
						mission.abilities[f.abilities[1]]--;
						prescence += 3;
					}

					List<string> slayer = new List<string>();
					List<string> friend = new List<string>();
					foreach (string trait in f.traits)
						switch (trait)
						{
							case "Beastslayer":
								slayer.Add("Beast");
								break;
							case "Brew Aficionado":
								friend.Add("Panda");
								break;
							case "Cave Dweller":
								slayer.Add("Underground");
								break;
							case "Child of Draenor":
								friend.Add("Orc");
								break;
							case "Cold - Blooded":
								slayer.Add("Snow");
								break;
							case "Combat Experience":
								prescence += 1;
								break;
							case "Death Fascination":
								friend.Add("Undead");
								break;
							case "Demonslayer":
								slayer.Add("Demon");
								break;
							case "Elvenkind":
								friend.Add("Blood elf");
								break;
							case "Furyslayer":
								slayer.Add("Fury");
								break;
							case "Gronnslayer":
								slayer.Add("Breaker");
								break;
							case "Guerilla Fighter":
								slayer.Add("Jungle");
								break;
							case "Mountaineer":
								slayer.Add("Mountain");
								break;
							case "Naturalist":
								slayer.Add("Forest");
								break;
							case "Ogreslayer":
								slayer.Add("Ogre");
								break;
							case "Orcslayer":
								slayer.Add("Orc");
								break;
							case "Plainsrunner":
								slayer.Add("Plains");
								break;
							case "Primalslayer":
								slayer.Add("Primal");
								break;
							case "Talonslayer":
								slayer.Add("Arrakoa");
								break;
							case "Voidslayer":
								slayer.Add("Void");
								break;
							case "Voodoo Zealot":
								friend.Add("Troll");
								break;
							case "Wastelander":
								slayer.Add("Desert");
								break;
						}
					foreach (string s in slayer)
						if (mission.trait == s)
							prescence += 1;
					foreach (string s in friend)
						foreach (Follower friend_maybe in follower_combo)
							if (friend_maybe != f && friend_maybe.race == s) prescence += 1.5;
				}

				//Set last value to larger of the two, your best and this run
				if(mission_results.Last() < prescence)
				{
					mission_results[mission_results.Count - 1] = prescence;
					input_mission.best_group = follower_combo;
					input_mission.prescence = prescence;
				}
			}

			if(input_mission.best_group != null)
				foreach (Follower follower in input_mission.best_group)
					follower.missions_completed++;
		}
	}

	class Program
	{
		static void Main(string[] args)
		{
			using (StreamWriter writer = new StreamWriter("WriteText.txt"))
			{
				List<Follower> followers = new List<Follower>();

				writer.Write(DateTime.Now);
				//Create all combinations of abilities (no repeats)
				writer.WriteLine("Getting all ability combinations...");
				const int ability_count = 9;
				for (int i = 0; i < ability_count; i++)
					for (int j = i + 1; j < ability_count; j++)
						followers.Add(new Follower() { abilities = new int[] { i, j } });

				//Create all combination of n followers
				writer.WriteLine("Getting all follower combinations...");
				List<List<Follower>> combination_followers = Subset(followers, 4);

				//Open excel document
				writer.WriteLine("Reading file...");
				FileStream stream = File.Open("missions.xlsx", FileMode.Open, FileAccess.Read);
				IExcelDataReader excel_reader = ExcelReaderFactory.CreateOpenXmlReader(stream);
				DataSet result = excel_reader.AsDataSet();

				//Read ability names in from excel document
				string[] ability_names = new string[ability_count];
				excel_reader.Read();
				for (int i=0; i<ability_count; i++)
					ability_names[i] = excel_reader.GetString(i);

				//Load in the list of missions from excel document
				List<Mission> missions = new List<Mission>();
				while (excel_reader.Read())
				{
					missions.Add(new Mission() { abilities = new int[ability_count] });
					for (int i = 0; i < ability_count; i++)
						missions.Last().abilities[i] = excel_reader.GetInt32(i);
					missions.Last().trait = excel_reader.GetString(ability_count);
				}

				//Setup groups
				writer.WriteLine("Setting up groups...");
				var groups = new List<Group>();
				foreach (List<Follower> group in combination_followers)
					groups.Add(new Group() { followers = group });

				foreach (Group g in groups) g.SetupCombos();

				//Run through each mission and record success rate
				writer.WriteLine("Testing groups...");
				foreach (Group group in groups)
					foreach (Mission mission in missions)
						group.RunMission(mission);

				//Sort by the success rates of the groups
				groups.Sort((a, b) => b.total_ability_used.CompareTo(a.total_ability_used));

				//Print out results for top groups
				const int counter = 100;
				for (int i = 0; i < Math.Min(groups.Count, counter); i++)
				{
					for (int j = 0; j < groups[i].followers.Count; j++)
						writer.WriteLine(String.Format("Follower #{0}: {1}, {2}", j + 1, ability_names[groups[i].followers[j].abilities[0]], ability_names[groups[i].followers[j].abilities[1]]));
					writer.WriteLine(String.Format("{0}/{1}\n", groups[i].total_ability_used, missions.Count * 6));
				}

				//Get the final results and print those out
				var follower_stats_overall = new Dictionary<int, Follower>();
				for (int i = 0; i < Math.Min(groups.Count, counter); i++)
				{
					foreach (Follower follower in groups[i].followers)
					{
						int calc_key = follower.abilities[0] * 100000000 + follower.abilities[1];
						if (follower_stats_overall.ContainsKey(calc_key))
							follower_stats_overall[calc_key].missions_completed += follower.missions_completed;
						else
							follower_stats_overall[calc_key] = new Follower() { abilities = follower.abilities, missions_completed = follower.missions_completed };
					}
				}

				List<Follower> followers_sorted = follower_stats_overall.Values.ToList();
				followers_sorted.Sort((a, b) => a.missions_completed.CompareTo(b.missions_completed));
				followers_sorted.Reverse();
				foreach (Follower follower_cumulative in followers_sorted)
					writer.WriteLine(String.Format("{0}: {1}, {2}", follower_cumulative.missions_completed, ability_names[follower_cumulative.abilities[0]], ability_names[follower_cumulative.abilities[1]]));

				//Open excel file for followers
				writer.WriteLine("Reading followers file...");
				stream = File.Open("followers.xlsx", FileMode.Open, FileAccess.Read);
				excel_reader = ExcelReaderFactory.CreateOpenXmlReader(stream);
				result = excel_reader.AsDataSet();

				//Skip first row
				excel_reader.Read();

				//Load in the list of followers
				var my_followers = new Group();
				var ability_names_list = new List<String>(ability_names);
				while (excel_reader.Read())
				{
					var f = new Follower();
					f.name = excel_reader.GetString(0);
					f.race = excel_reader.GetString(1);

					f.abilities[0] = ability_names_list.IndexOf(excel_reader.GetString(2));
					f.abilities[1] = ability_names_list.IndexOf(excel_reader.GetString(3));

					f.traits[0] = excel_reader.GetString(4);
					f.traits[1] = excel_reader.GetString(5);
					f.traits[2] = excel_reader.GetString(6);

					my_followers.followers.Add(f);
				}

				//Run followers on missions
				foreach (Mission mission in missions)
				{
					my_followers.RunMission(mission);
					writer.Write((mission.prescence) + ": ");
					foreach (Follower f in mission.best_group)
					{
						writer.Write(f.name + "/");
						f.missions_completed_2++;
					}
					writer.Write("\n");
				}

				//Print follower information out
				my_followers.followers.Sort((a, b) => a.missions_completed_2.CompareTo(b.missions_completed_2));
				my_followers.followers.Reverse();
				foreach (Follower follower in my_followers.followers)
					writer.WriteLine(String.Format("{0}: {1}", follower.name, follower.missions_completed_2));

				//Find mr perfect amongst all the possible combinations
				List<double> mission_prescence = new List<double>();
				double sum_prescence = missions.Sum(item => item.prescence);
				List<Tuple<Follower, double>> track_best_followers = new List<Tuple<Follower, double>>();
				foreach (Follower follower in followers)
				{
					Group one_more = new Group();
					one_more.followers = new List<Follower>(my_followers.followers);
					one_more.followers.Add(follower);
					foreach (Mission mission in missions)
						one_more.RunMission(mission);
					double new_sum = missions.Sum(item => item.prescence);
					double diff = new_sum - sum_prescence;
					track_best_followers.Add(new Tuple<Follower, double>(follower, diff));
				}

				track_best_followers.Sort((a, b) => a.Item2.CompareTo(b.Item2));
				track_best_followers.Reverse();

				foreach(Tuple<Follower, double> t in track_best_followers)
					writer.WriteLine(string.Format("{0}/{1} -- +{2}", ability_names[t.Item1.abilities[0]], ability_names[t.Item1.abilities[1]], t.Item2));
			}
		}

		//Create a combination of passed in length and given choices
		public static List<List<Follower>> Subset(List<Follower> choices, int remaining_times)
		{
			var result = new List<List<Follower>>();
			for (int i = 0; i < choices.Count; i++)
			{
				if (remaining_times > 1)
				{
					List<List<Follower>> next_sequences = Subset(choices.GetRange(i, choices.Count - i), remaining_times - 1);
					for (int j = 0; j < next_sequences.Count; j++)
					{
						result.Add(new List<Follower>() { choices[i] });
						result.Last().AddRange(next_sequences[j]);
					}
				}
				else //last run is simply each choice
					result.Add(new List<Follower>() { choices[i] });
			}
			return result;
		}
	}
}