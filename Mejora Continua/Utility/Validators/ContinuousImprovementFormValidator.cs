using FluentValidation;
using Mejora_Continua.Models;

namespace Mejora_Continua.Utility.Validators
{
    public class ContinuousImprovementFormValidator : AbstractValidator<ContinuousImprovementIdeas>
    {
        public ContinuousImprovementFormValidator()
        {
            RuleFor(x => x.FullName)
                .NotEmpty().WithMessage("El nombre es obligatorio");

            RuleFor(x => x.WorkArea)
                .NotEmpty().WithMessage("El área de trabajo es obligatoria");

            RuleFor(x => x.CurrentSituation)
                .NotEmpty().WithMessage("La situación actual es obligatoria");

            RuleFor(x => x.IdeaDescription)
                .NotEmpty().WithMessage("La idea de mejora es obligatoria");

            //RuleFor(x => x.CategoryIds)
            //    .NotNull().WithMessage("Debe seleccionar al menos una categoría")
            //    .Must(catogories => catogories != null && catogories.Count > 0)
            //    .WithMessage("Debe seleccionar al menos una cattegoría");
        }
    }
}