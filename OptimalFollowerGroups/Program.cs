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
		public int[] abilities = new int[3];

		public string race;

		public string[] race_friends = new string[0];
		public string[] trait_counter = new string[0];

		public bool combat_experience = false;
		public bool high_stamina = false;
		public bool burst_of_power = false;
		public bool epic_mount = false;

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
					foreach(int ability in f.abilities)
					{
						if(ability >= 0 && mission.abilities[ability] != 0)
						{
							f.missions_completed++;
							mission.abilities[ability]--;
							prescence += 3;
						}
					}

					//apply traits to missions
					foreach (string s in f.trait_counter)
						if (mission.trait == s)
							prescence += 1;
					foreach (string s in f.race_friends)
						foreach (Follower friend_maybe in follower_combo)
							if (friend_maybe != f && friend_maybe.race == s)
							{
								prescence += 1.5;
								break;
							}
					if (f.combat_experience)
						prescence++;
					if (f.burst_of_power || f.high_stamina)
					{
						bool mounted = false;
						foreach (Follower friend_maybe in follower_combo)
							if (friend_maybe.epic_mount == true)
								mounted = true;
						if (!mounted && f.high_stamina) prescence++;
						if (mounted && f.burst_of_power) prescence++;
					}
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
				const int ability_count = 9;
				writer.Write(DateTime.Now);

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
					int c = 0;
					var f = new Follower();
					f.name = excel_reader.GetString(c++);
					f.race = excel_reader.GetString(c++);

					f.abilities[0] = ability_names_list.IndexOf(excel_reader.GetString(c++));
					f.abilities[1] = ability_names_list.IndexOf(excel_reader.GetString(c++));
					f.abilities[2] = ability_names_list.IndexOf(excel_reader.GetString(c++));

					string race = excel_reader.GetString(c++);
					if(!string.IsNullOrWhiteSpace(race))
						f.race_friends = race.Split(new string[] { ", " }, StringSplitOptions.None);

					string traits = excel_reader.GetString(c++);
					if (!string.IsNullOrWhiteSpace(traits))
						f.trait_counter = traits.Split(new string[] { ", " }, StringSplitOptions.None);

					f.combat_experience = !string.IsNullOrWhiteSpace(excel_reader.GetString(c++));
					f.high_stamina = !string.IsNullOrWhiteSpace(excel_reader.GetString(c++));
					f.burst_of_power = !string.IsNullOrWhiteSpace(excel_reader.GetString(c++));
					f.epic_mount = !string.IsNullOrWhiteSpace(excel_reader.GetString(c++));

					my_followers.followers.Add(f);
				}

				//Run followers on missions
				writer.WriteLine("\nBest followers for each mission");
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
				writer.WriteLine("\nNumber of missions each follower went on");
				my_followers.followers.Sort((a, b) => a.missions_completed_2.CompareTo(b.missions_completed_2));
				my_followers.followers.Reverse();
				foreach (Follower follower in my_followers.followers)
					writer.WriteLine(String.Format("{0}: {1}", follower.name, follower.missions_completed_2));

				//Find mr perfect amongst all the possible combinations
				//Create all combinations of abilities (no repeats)
				writer.WriteLine("\nLargest improvement combos");
				List<Follower> followers = new List<Follower>();
				for (int i = 0; i < ability_count; i++)
					for (int j = i + 1; j < ability_count; j++)
						followers.Add(new Follower() { abilities = new int[] { i, j } });

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