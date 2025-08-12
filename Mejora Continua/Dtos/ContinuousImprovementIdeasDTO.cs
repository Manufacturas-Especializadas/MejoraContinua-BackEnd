namespace Mejora_Continua.Dtos
{
    public class ContinuousImprovementIdeasDTO
    {
        public string FullName { get; set; }
        public string WorkArea { get; set; }
        public string CurrentSituation { get; set; }
        public string IdeaDescription { get; set; }
        public int StatusId { get; set; }
        public DateTime? RegistrationDate { get; set; }
        public List<int> CategoryIds { get; set; } = new List<int>();
        public List<int> ChampionIds { get; set; } = new List<int>();
        public List<string> Names { get; set; } = new List<string>();

    }
}