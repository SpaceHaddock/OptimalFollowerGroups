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
		public int a = 0;
		public int b = 0;

		public double missions_completed = 0;
	}

	//Class which represents a group of followers
	//Can be any size and calculates results for missions
	public class Group
	{
		public List<Follower> followers = new List<Follower>();
		public List<int> mission_results = new List<int>();
		public int total_ability_used { get { return mission_results.Sum(); } }

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
		public void RunMission(int[] input_mission)
		{
			mission_results.Add(0);
			//Go through each combination of followers in this group
			foreach (Follower[] follower_combo in combinations_followers)
			{
				var mission = (int[]) input_mission.Clone();
				foreach (Follower f in follower_combo)
				{
					if(mission[f.a] != 0)
					{
						f.missions_completed++;
						mission[f.a]--;
					}
					if(mission[f.b] != 0)
					{
						f.missions_completed++;
						mission[f.b]--;
					}
				}

				//Set last value to larger of the two, your best and this run
				mission_results[mission_results.Count - 1] = Math.Max(mission_results.Last(), input_mission.Sum() - mission.Sum());
			}
		}
	}

	class Program
	{
		public static string[] ability_names = new string[]
		{ "Danger Zones", "Massive Strike", "Magic Debuff", "Timed Battle",
			"Deadly Minions", "Powerful Spell", "Minion Swarms", "Group Damage", "Wild Aggression"};

		static void Main(string[] args)
		{
			using (StreamWriter writer = new StreamWriter("WriteText.txt"))
			{
				List<Follower> followers = new List<Follower>();

				//Create all combinations of abilities (no repeats)
				writer.WriteLine("Getting all ability combinations...");
				const int ability_count = 9;
				for (int i = 0; i < ability_count; i++)
					for (int j = i + 1; j < ability_count; j++)
						followers.Add(new Follower() { a = i, b = j });

				//Create all combination of n followers
				writer.WriteLine("Getting all follower combinations...");
				List<List<Follower>> combination_followers = Subset(followers, 4);

				//Load in the list of missions
				List<int[]> missions = new List<int[]>();

				writer.WriteLine("Reading file...");
				FileStream stream = File.Open("missions.xlsx", FileMode.Open, FileAccess.Read);
				IExcelDataReader excel_reader = ExcelReaderFactory.CreateOpenXmlReader(stream);
				DataSet result = excel_reader.AsDataSet();
				while (excel_reader.Read())
				{
					missions.Add(new int[ability_count]);
					for (int i = 0; i < ability_count; i++)
						missions.Last()[i] = excel_reader.GetInt32(i);
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
					foreach (int[] mission in missions)
						group.RunMission(mission);

				//Sort by the success rates of the groups
				groups.Sort((a, b) => b.total_ability_used.CompareTo(a.total_ability_used));

				//Print out results for top groups
				const int counter = 100;
				for (int i = 0; i < Math.Min(groups.Count, counter); i++)
				{
					for (int j = 0; j < groups[i].followers.Count; j++)
						writer.WriteLine(String.Format("Follower #{0}: {1}, {2}", j + 1, ability_names[groups[i].followers[j].a], ability_names[groups[i].followers[j].b]));
					writer.WriteLine(String.Format("{0}/{1}\n", groups[i].total_ability_used, missions.Count * 6));
				}

				//Get the final results and print those out
				var follower_stats_overall = new Dictionary<int, Follower>();
				for (int i = 0; i < Math.Min(groups.Count, counter); i++)
				{
					foreach (Follower follower in groups[i].followers)
					{
						int calc_key = follower.a * 100000000 + follower.b;
						if (follower_stats_overall.ContainsKey(calc_key))
							follower_stats_overall[calc_key].missions_completed += follower.missions_completed;
						else
							follower_stats_overall[calc_key] = new Follower() { a = follower.a, b = follower.b, missions_completed = follower.missions_completed };
					}
				}

				List<Follower> followers_sorted = follower_stats_overall.Values.ToList();
				followers_sorted.Sort((a, b) => a.missions_completed.CompareTo(b.missions_completed));
				followers_sorted.Reverse();
				foreach (Follower follower_cumulative in followers_sorted)
				{
					writer.WriteLine(String.Format("{0}: {1}, {2}", follower_cumulative.missions_completed, ability_names[follower_cumulative.a], ability_names[follower_cumulative.b]));
				}
			}
		}

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