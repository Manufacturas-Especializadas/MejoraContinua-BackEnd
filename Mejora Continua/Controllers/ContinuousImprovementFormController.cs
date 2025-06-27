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
            var list = await _context.ContinuousImprovementIdeas
                            .Select(i => new
                            {
                                i.Id,
                                i.FullName,
                                i.WorkArea,
                                i.RegistrationDate,
                                i.CurrentSituation,
                                i.IdeaDescription,
                                status = i.Status.Name,
                                championNames = i.Champion.Select(c => c.Name).ToList(),
                            })
                            .AsNoTracking()
                            .ToListAsync();

            if (list == null)
            {
                throw new Exception("No hay datos disponibles");
            }

            return Ok(list);
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
            var idea = await _context.ContinuousImprovementIdeas.Select(i => new
            {
                i.Id,
                i.FullName,
                i.WorkArea,
                i.RegistrationDate,
                i.CurrentSituation,
                i.IdeaDescription,
                status = i.Status.Name,
                championNames = i.Champion.Select(c => c.Name).ToList(),
            })
            .AsNoTracking()
            .ToListAsync();

            using (var workBook = new XLWorkbook())
            {
                var workSheet = workBook.Worksheets.Add("ContinuousImprovementIdeas");

                workSheet.Cell(1, 1).Value = "Nombre completo";
                workSheet.Cell(1, 2).Value = "Area de trabajo";
                workSheet.Cell(1, 3).Value = "Situacion";
                workSheet.Cell(1, 4).Value = "Descripcion";
                workSheet.Cell(1, 5).Value = "Estado";
                workSheet.Cell(1, 6).Value = "Fecha de registro";
                workSheet.Cell(1, 7).Value = "Champion(s)";

                for(int i = 0; i < idea.Count; i++)
                {
                    workSheet.Cell(i + 2, 1).Value = idea[i].FullName;
                    workSheet.Cell(i + 2, 2).Value = idea[i].WorkArea;
                    workSheet.Cell(i + 2, 3).Value = idea[i].CurrentSituation;
                    workSheet.Cell(i + 2, 4).Value = idea[i].IdeaDescription;
                    workSheet.Cell(i + 2, 5).Value = idea[i].status;
                    workSheet.Cell(i + 2, 6).Value = idea[i].RegistrationDate!.Value.ToString("dd/MM/yyyy");
                    workSheet.Cell(i + 2, 7).Value = string.Join(", ", idea[i].championNames);
                }

                var stream = new MemoryStream();
                workBook.SaveAs(stream);
                stream.Position = 0;

                var filName = $"DatosExportados_{DateTime.Now:ddMMyyyy}.xlsx";

                return File(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", filName);
            }
        }


        [HttpPost("Register")]
        public async Task<IActionResult> Register([FromBody] ContinuousImprovementIdeas improvementIdea)
        {
            if (!ModelState.IsValid)
                return BadRequest(ModelState);

            try
            {
                if (string.IsNullOrEmpty(improvementIdea.FullName))
                {
                    improvementIdea.FullName = string.Join(", ", improvementIdea.Names ?? new List<string>());
                }

                _context.ContinuousImprovementIdeas.Add(improvementIdea);
                await _context.SaveChangesAsync();

                var categoryList = await _context.ContinuousImprovementCategory
                    .Where(c => improvementIdea.CategoryIds.Contains(c.Id))
                    .ToListAsync();

                foreach (var categoryId in improvementIdea.CategoryIds)
                {
                    var category = categoryList.FirstOrDefault(c => c.Id == categoryId);
                    if (category != null)
                    {
                        improvementIdea.Category.Add(category);
                    }
                }

                await _context.SaveChangesAsync();

                return Ok(new { message = "Registro exitoso" });
            }
            catch (Exception ex)
            {
                return StatusCode(StatusCodes.Status500InternalServerError, new
                {
                    message = $"Error interno del servidor: {ex.Message}",
                    innerException = ex.InnerException?.Message,
                    stackTrace = ex.StackTrace
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