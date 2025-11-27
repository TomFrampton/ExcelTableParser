using System.ComponentModel.DataAnnotations;
using System.Xml.Linq;

namespace ExcelTableParser.Console
{
    internal class Program
    {
        static void Main(string[] args)
        {
            using var fileStream = File.OpenRead("D:\\Projects\\ExcelTableParser\\ExcelTableParser\\ExcelTableParser.Console\\NDS Modifiers.xlsm");

            var modifier = ExcelTableParser.ParseTable<Modifier>(
                fileStream,
                "Modifier",
                "t_modifier");

            var modifierSetResult = ExcelTableParser.ParseTable<ModifierSet>(
                fileStream,
                "Modifier Set",
                "t_modifier_set");

            var modifierSetMemberResult = ExcelTableParser.ParseTable<ModifierSetMember>(
                fileStream,
                "Modifier Set Member",
                "t_modifier_set_member",
                null,
                (modifierSetMember, rowNumber) => !modifierSetResult.Items.Any(x => x.ModifierSetName == modifierSetMember.ModifierSetName)
                    ? [$"Modifier set '{modifierSetMember.ModifierSetName}' not found."]
                    : []
                );

            var scenarioSetResult = ExcelTableParser.ParseTable<ScenarioSet>(
                fileStream,
                "Scenario Set",
                "t_scenario_set");

            var scenarioSetMemberResult = ExcelTableParser.ParseTable<ScenarioSetMember>(
                fileStream,
                "Scenario Set Member",
                "t_scenario_set_member",
                null,
                (scenarioSetMember, rowNumber) => !scenarioSetResult.Items.Any(x => x.ScenarioSetName == scenarioSetMember.ScenarioSetName)
                    ? [$"Scenario set '{scenarioSetMember.ScenarioSetName}' not found."]
                    : []
                );
        }
    }
}
