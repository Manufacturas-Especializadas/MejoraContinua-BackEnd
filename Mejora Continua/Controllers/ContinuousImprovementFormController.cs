using ClosedXML.Excel;
using Mejora_Continua.Dtos;
using Mejora_Continua.Models;
using Mejora_Continua.Services;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using System.Security.AccessControl;

namespace Mejora_Continua.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ContinuousImprovementFormController : ControllerBase
    {
        private AppDbContext _context;
        private readonly EmailService _emailService;

        public ContinuousImprovementFormController(AppDbContext context, EmailService emailService)
        {
            _context = context;
            _emailService = emailService;
        }

        [HttpGet("GetIdeaById")]
        public async Task<IActionResult> GetIdeaByIdeaAsync(int id)
        {
            var ideId = await _context.ContinuousImprovementIdeas.FirstOrDefaultAsync(i =>  i.Id == id);

            if(ideId == null)
            {
                return NotFound(new { message = "Id no encontrado" });
            }

            return Ok(ideId);
        }

        [HttpGet("GetListIdeas")]
        public async Task<IActionResult> GetListIdeasAsync()
        {
            try
            {
                var list = await _context.ContinuousImprovementIdeas
                    .OrderByDescending(i => i.Id)
                    .Select(i => new
                    {
                        i.Id,
                        i.FullName,
                        i.WorkArea,
                        i.RegistrationDate,
                        i.CurrentSituation,
                        i.IdeaDescription,
                        Status = i.Status != null ? i.Status.Name : "Sin estado",
                        ChampionNames = _context.IdeaChampion
                            .Where(ic => ic.IdeaId == i.Id)
                            .Select(ic => ic.Champion != null ? ic.Champion.Name : "")
                            .ToList(),
                        Categories = _context.IdeaCategory
                            .Where(ic => ic.IdeaId == i.Id)
                            .Select(ic => ic.Category != null ? ic.Category.Name : "")
                            .ToList()
                    })
                    .AsNoTracking()
                    .ToListAsync();

                if (list == null || !list.Any())
                {
                    return NotFound(new { message = "No hay ideas registradas" });
                }

                return Ok(list);
            }
            catch (Exception ex)
            {
                return StatusCode(StatusCodes.Status500InternalServerError, new
                {
                    message = "Error al obtener la lista de ideas",
                    error = ex.Message,
                    stackTrace = ex.StackTrace
                });
            }
        }

        [HttpGet("GetListStatus")]
        public async Task<List<ContinuousImprovementStatus>> GetListStatusAsync()
        {
            return await _context.ContinuousImprovementStatus.AsNoTracking().ToListAsync();

        }

        [HttpGet("GetListChampions")]
        public async Task<List<ChampionDTO>> GetListChampionsAsync()
        {
            var list = await _context.ContinuousImprovementChampions
            .AsNoTracking()
            .Select(x => new ChampionDTO { Id = x.Id, Name = x.Name })
            .ToListAsync();

            if (!list.Any())
                throw new Exception("No hay datos disponibles");

            return list;
        }

        [HttpPost("DownloadExcel")]
        public async Task<IActionResult> ExportToExcel()
        {
            var ideas = await _context.ContinuousImprovementIdeas
                .Include(i => i.Status)
                .Include(i => i.IdeaChampion)
                    .ThenInclude(ic => ic.Champion)
                .Select(i => new
                {
                    i.Id,
                    i.FullName,
                    i.WorkArea,
                    i.RegistrationDate,
                    i.CurrentSituation,
                    i.IdeaDescription,
                    StatusName = i.Status.Name,
                    ChampionNames = i.IdeaChampion.Select(ic => ic.Champion.Name).ToList()
                })
                .AsNoTracking()
                .ToListAsync();

            using (var workBook = new XLWorkbook())
            {
                var workSheet = workBook.Worksheets.Add("ContinuousImprovementIdeas");

                workSheet.Cell(1, 1).Value = "ID";
                workSheet.Cell(1, 2).Value = "Nombre completo";
                workSheet.Cell(1, 3).Value = "Área de trabajo";
                workSheet.Cell(1, 4).Value = "Situación actual";
                workSheet.Cell(1, 5).Value = "Descripción de la idea";
                workSheet.Cell(1, 6).Value = "Estado";
                workSheet.Cell(1, 7).Value = "Fecha de registro";
                workSheet.Cell(1, 8).Value = "Champion(s)";

                for (int i = 0; i < ideas.Count; i++)
                {
                    var current = ideas[i];
                    workSheet.Cell(i + 2, 1).Value = current.Id;
                    workSheet.Cell(i + 2, 2).Value = current.FullName;
                    workSheet.Cell(i + 2, 3).Value = current.WorkArea;
                    workSheet.Cell(i + 2, 4).Value = current.CurrentSituation;
                    workSheet.Cell(i + 2, 5).Value = current.IdeaDescription;
                    workSheet.Cell(i + 2, 6).Value = current.StatusName;
                    workSheet.Cell(i + 2, 7).Value = current.RegistrationDate?.ToString("dd/MM/yyyy");
                    workSheet.Cell(i + 2, 8).Value = string.Join(", ", current.ChampionNames);
                }

                workSheet.Columns().AdjustToContents();

                var stream = new MemoryStream();
                workBook.SaveAs(stream);
                stream.Position = 0;

                var fileName = $"IdeasMejoraContinua_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx";

                return File(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
            }
        }


        [HttpPost("Register")]
        public async Task<IActionResult> Register([FromBody] ContinuousImprovementIdeasDTO ideaDto)
        {
            if (!ModelState.IsValid)
                return BadRequest(ModelState);

            try
            {
                if (!await _context.ContinuousImprovementStatus.AnyAsync(s => s.Id == ideaDto.StatusId))
                {
                    return BadRequest("El estado especificado no existe");
                }

                var existingCategoryIds = await _context.ContinuousImprovementCategory
                    .Where(c => ideaDto.CategoryIds.Contains(c.Id))
                    .Select(c => c.Id)
                    .ToListAsync();

                var invalidCategoryIds = ideaDto.CategoryIds.Except(existingCategoryIds).ToList();
                if (invalidCategoryIds.Any())
                {
                    return BadRequest($"Las siguientes categorías no existen: {string.Join(", ", invalidCategoryIds)}");
                }

                if (ideaDto.ChampionIds != null && ideaDto.ChampionIds.Any())
                {
                    var existingChampionIds = await _context.ContinuousImprovementChampions
                        .Where(c => ideaDto.ChampionIds.Contains(c.Id))
                        .Select(c => c.Id)
                        .ToListAsync();

                    var invalidChampionIds = ideaDto.ChampionIds.Except(existingChampionIds).ToList();
                    if (invalidChampionIds.Any())
                    {
                        return BadRequest($"Los siguientes champions no existen: {string.Join(", ", invalidChampionIds)}");
                    }
                }

                var idea = new ContinuousImprovementIdeas
                {
                    FullName = string.IsNullOrEmpty(ideaDto.FullName) ?
                             string.Join(", ", ideaDto.Names ?? new List<string>()) :
                             ideaDto.FullName,
                    WorkArea = ideaDto.WorkArea,
                    CurrentSituation = ideaDto.CurrentSituation,
                    IdeaDescription = ideaDto.IdeaDescription,
                    StatusId = ideaDto.StatusId,
                    RegistrationDate = ideaDto.RegistrationDate ?? DateTime.Now
                };

                _context.ContinuousImprovementIdeas.Add(idea);
                await _context.SaveChangesAsync();

                foreach (var categoryId in ideaDto.CategoryIds)
                {
                    _context.IdeaCategory.Add(new IdeaCategory
                    {
                        IdeaId = idea.Id,
                        CategoryId = categoryId
                    });
                }

                if (ideaDto.ChampionIds != null)
                {
                    foreach (var championId in ideaDto.ChampionIds)
                    {
                        _context.IdeaChampion.Add(new IdeaChampion
                        {
                            IdeaId = idea.Id,
                            ChampionId = championId
                        });
                    }
                }

                await _context.SaveChangesAsync();

                return Ok(new
                {
                    message = "Registro exitoso",
                    ideaId = idea.Id,
                    selectedCategories = ideaDto.CategoryIds.Count,
                    selectedChampions = ideaDto.ChampionIds?.Count ?? 0
                });
            }
            catch (Exception ex)
            {
                return StatusCode(StatusCodes.Status500InternalServerError, new
                {
                    message = "Error al registrar la idea",
                    error = ex.Message,
                    innerException = ex.InnerException?.Message
                });
            }
        }

        [HttpPost("AssignChampionsToIdea")]
        public async Task<IActionResult> AssignChampionsToIdea([FromBody] List<AssignChampionDTO> dtos)
        {
            if (!ModelState.IsValid)
                return BadRequest(ModelState);

            try
            {
                foreach (var dto in dtos)
                {

                   var idea = await _context.ContinuousImprovementIdeas
                            .FirstOrDefaultAsync(i => i.Id == dto.IdeaId);

                    var champion = await _context.ContinuousImprovementChampions
                            .FirstOrDefaultAsync(c => c.Id == dto.ChampionId);

                    if(idea == null || champion == null)
                    {
                        continue;
                    }

                    var exists = await _context.Set<Dictionary<string, object>>("IdeaChampion")
                        .AnyAsync(e =>
                            (int)e["IdeaId"] == dto.IdeaId &&
                            (int)e["ChampionId"] == dto.ChampionId
                        );

                    if(!exists)
                    {
                        var entry = new Dictionary<string, object>
                        {
                            ["IdeaId"] = dto.IdeaId,
                            ["ChampionId"] = dto.ChampionId
                        };

                        await _context.Set<Dictionary<string, object>>("IdeaChampion").AddAsync(entry);
                    }

                    await SendChampionAssignedEmail(champion, idea);
                }

                await _context.SaveChangesAsync();

                return Ok(new { message = "Champions asignados correctamente." });
            }
            catch (Exception ex)
            {
                return StatusCode(500, new
                {
                    error = "Error al asignar los champions.",
                    details = ex.Message,
                    stackTrace = ex.StackTrace
                });
            }
        }

        private async Task SendChampionAssignedEmail(ContinuousImprovementChampions champion, ContinuousImprovementIdeas idea)
        {
            var subject = "Se te ha asignado una nueva idea";

            var body = $@"
                    <h2> Hola {champion.Name} </h2>
                    <p> Se te ha asignado una nueva idea </p>
                    <ul>
                        <li><strong>Nombre:</strong>{idea.FullName}</li>
                        <li><strong>Situación actual:</strong>{idea.CurrentSituation}</li>
                        <li><strong>Descripción de la idea:</strong>{idea.IdeaDescription}</li>
                    </ul>
                    <p>Favor de atenderla lo más pronto posible</p>
                    <br/>
                    <p>Saludos, <br/>Equipo de Mejora Continua</p>
                ";

            await _emailService.SendEmailAsync(champion.Email, subject, body);
        }

        [HttpPut("Update/{id:int}")]
        public async Task<IActionResult> Update([FromBody] ContinuousImprovementIdeasDTO ideasDTO, int id)
        {
            var idea = await _context.ContinuousImprovementIdeas.FindAsync(id);
            if (idea == null) return BadRequest("No se encontró la idea");

            idea.FullName = ideasDTO.FullName;
            idea.WorkArea = ideasDTO.WorkArea;
            idea.StatusId = ideasDTO.StatusId;
            idea.CurrentSituation = ideasDTO.CurrentSituation;
            idea.IdeaDescription = ideasDTO.IdeaDescription;

            await _context.SaveChangesAsync();

            return Ok(new { message = "Editado correctamente" });
        }

        [HttpDelete("Delete/{id}")]
        public async Task<IActionResult> Delete(int id)
        {
            var ideaDb = await _context.ContinuousImprovementIdeas.FirstOrDefaultAsync(i => i.Id == id);

            if (ideaDb == null) return NotFound("No se encontró el id seleccionado");

            _context.ContinuousImprovementIdeas.Remove(ideaDb);
            await _context.SaveChangesAsync();

            return Ok(new { message = "Registro eliminado" });
        }
    }

}