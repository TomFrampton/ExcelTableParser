using DocumentFormat.OpenXml.VariantTypes;
using DocumentFormat.OpenXml.Wordprocessing;
using System.ComponentModel.DataAnnotations;

namespace ExcelTableParser.Console
{
    public class ScenarioSet
    {
        [Column("Scenario Set Name")]
        [Required]
        public string? ScenarioSetName { get; set; }

        [Column("Scenario Set Description")]
        [Required]
        public string? ScenarioSetDescription { get; set; }
    }

    public class ScenarioSetMember
    {
        [Column("Scenario Set Name")]
        [Required]
        public string? ScenarioSetName { get; set; }

        [Column("Modifier Set Name")]
        [Required]
        public string? ModifierSetName { get; set; }

        [Column("Sequence")]
        [Required]
        public int? Sequence { get; set; }
    }

    public class ModifierSet
    {
        [Column("Modifier Set Name")]
        [Required]
        public string? ModifierSetName { get; set; }

        [Column("Type")]
        [Required]
        public string? Type { get; set; }

        [Column("Modifier Set Description")]
        public string? ModifierSetDescription { get; set; }
    }

    public class ModifierSetMember
    {
        [Column("Modifier Set Name")]
        [Required]
        public string? ModifierSetName { get; set; }

        [Column("Modifier Name")]
        [Required]
        public string? ModifierName { get; set; }

        [Column("Sequence")]
        [Required]
        public int? Sequence { get; set; }
    }

    public class Modifier
    {
        [Column("Modifier Name")]
        [Required]
        public string? ModifierName { get; set; }

        [Column("Modifier Description")]
        [Required]
        public string? ModifierDescription { get; set; }

        [Column("Type")]
        [Required]
        public string? Type { get; set; }

        [Column("Domain")]
        [Required]
        public string? Domain { get; set; }

        [Column("Is Enabled")]
        [Required]
        public int? IsEnabled { get; set; }

        [Column("Term")]
        [Required]
        public string? Term { get; set; }

        [Column("Operation")]
        [Required]
        public string? Operation { get; set; }

        [Column("Value")]
        [Required]
        public string? Value { get; set; }

        [Column("Program Name")]
        [Required]
        public string? ProgramName { get; set; }

        [Column("Layer Reference")]
        [Required]
        public string? LayerReference { get; set; }

        [Column("Event Name")]
        [Required]
        public string? EventName  { get; set; }

        [Column("Class Name")]
        [Required]
        public string? ClassName { get; set; }

        [Column("YOA")]
        [Required]
        public string? YOA { get; set; }

        [Column("Currency")]
        [Required]
        public string? Currency { get; set; }

        [Column("ModifierJSON")]
        [Required]
        public string? ModifierJSON { get; set; }

        [Column("Flow State")]
        [Required]
        public string? FlowState { get; set; }
    }
}
