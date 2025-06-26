#nullable disable
using System;
using System.Collections.Generic;

namespace Mejora_Continua.Models;

public partial class ContinuousImprovementChampions
{
    public int Id { get; set; }

    public string Name { get; set; }

    public string Email { get; set; }

    public virtual ICollection<ContinuousImprovementIdeas> Idea { get; set; } = new List<ContinuousImprovementIdeas>();
}