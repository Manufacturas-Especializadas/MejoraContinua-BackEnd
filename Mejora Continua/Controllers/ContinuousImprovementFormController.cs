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

        private const string AdminEmail = "anahi.sauceda@mesa.ms";

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
                    .Where(i => i.Status != null)
                    .OrderByDescending(i => i.Id)
                    .Select(i => new
                    {
                        i.Id,
                        i.FullName,
                        i.WorkArea,
                        i.RegistrationDate,
                        i.CurrentSituation,
                        i.IdeaDescription,
                        Status = i.Status.Name,
                        Year = i.RegistrationDate!.Value.Year,
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
        public async Task<IActionResult> ExportToExcel([FromQuery] int year)
        {
            var ideasQuery = _context.ContinuousImprovementIdeas
                            .Include(i => i.Status)
                            .Include(i => i.IdeaChampion)
                                .ThenInclude(ic => ic.Champion)
                            .AsQueryable();

            if(year > 0)
            {
                ideasQuery = ideasQuery.Where(i => i.RegistrationDate.HasValue && i.RegistrationDate.Value.Year == year);
            }

            var ideas = await ideasQuery
                        .Select(i => new
                        {
                            i.Id,
                            i.FullName,
                            i.WorkArea,
                            i.RegistrationDate,
                            i.CurrentSituation,
                            i.IdeaDescription,
                            StatusName = i.Status.Name,
                            Champions = i.IdeaChampion.Select(ic => ic.Champion.Name).ToList(),
                            Categorys = i.IdeaCategory.Select(ica => ica.Category.Name).ToList(),
                        })
                        .AsNoTracking()
                        .ToListAsync();

            using (var workBook = new XLWorkbook())
            {
                var workSheet = workBook.Worksheets.Add($"Reporte {year}");

                var headerColor = XLColor.FromHtml("#0284c7");
                var whiteColor = XLColor.White;

                workSheet.Cell(1, 1).Value = $"REPORTE DE IDEAS DE MEJORA - AÑO {year}";
                workSheet.Range(1, 1, 1, 9).Merge();
                workSheet.Cell(1, 1).Style.Font.FontSize = 14;
                workSheet.Cell(1, 1).Style.Font.Bold = true;
                workSheet.Cell(1, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                workSheet.Cell(1, 1).Style.Fill.BackgroundColor = XLColor.LightGray;

                int headerRowIndex = 2;

                workSheet.Cell(headerRowIndex, 1).Value = "ID";
                workSheet.Cell(headerRowIndex, 2).Value = "Nombre completo";
                workSheet.Cell(headerRowIndex, 3).Value = "Área";
                workSheet.Cell(headerRowIndex, 4).Value = "Situación actual";
                workSheet.Cell(headerRowIndex, 5).Value = "Descripción";
                workSheet.Cell(headerRowIndex, 6).Value = "Estatus";
                workSheet.Cell(headerRowIndex, 7).Value = "Fecha de registro";
                workSheet.Cell(headerRowIndex, 8).Value = "Champions";
                workSheet.Cell(headerRowIndex, 9).Value = "Categorías";

                var headerRange = workSheet.Range(headerRowIndex, 1, headerRowIndex, 9);
                headerRange.Style.Fill.BackgroundColor = headerColor;
                headerRange.Style.Font.FontColor = whiteColor;
                headerRange.Style.Font.Bold = true;
                headerRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                int datRowIndex = 3;
                for(int i = 0; i < ideas.Count; i++)
                {
                    var current = ideas[i];
                    int currentRow = datRowIndex + i;

                    workSheet.Cell(currentRow, 1).Value = current.Id;
                    workSheet.Cell(currentRow, 2).Value = current.FullName;
                    workSheet.Cell(currentRow, 3).Value = current.WorkArea;
                    workSheet.Cell(currentRow, 4).Value = current.CurrentSituation;
                    workSheet.Cell(currentRow, 5).Value = current.IdeaDescription;

                    var statusCell = workSheet.Cell(currentRow, 6);
                    statusCell.Value = current.StatusName;
                    if (current.StatusName == "Aprobada") statusCell.Style.Font.FontColor = XLColor.Green;

                    workSheet.Cell(currentRow, 7).Value = current.RegistrationDate?.ToString("dd/MM/yyyy");
                    workSheet.Cell(currentRow, 8).Value = string.Join(", ", current.Champions);
                    workSheet.Cell(currentRow, 9).Value = string.Join(", ", current.Categorys);
                }

                workSheet.Columns().AdjustToContents();

                workSheet.Column(4).Width = 40;
                workSheet.Column(5).Width = 50;
                workSheet.Column(4).Style.Alignment.WrapText = true;
                workSheet.Column(5).Style.Alignment.WrapText = true;

                var tableRange = workSheet.Range(headerRowIndex, 1, datRowIndex + ideas.Count - 1, 9);
                tableRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                tableRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;

                var stream = new MemoryStream();
                workBook.SaveAs(stream);
                stream.Position = 0;

                var fileName = $"IdeasMejora_{year}_{DateTime.Now:ddMMyyyy}";

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

                await SendNewIdeaRegisteredEmail(idea);

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

                    var exists = await _context.IdeaChampion
                            .AnyAsync(e => e.IdeaId == dto.IdeaId && e.ChampionId == dto.ChampionId);

                    if (!exists)
                    {
                        var ideaChampion = new IdeaChampion
                        {
                            IdeaId = dto.IdeaId,
                            ChampionId = dto.ChampionId
                        };

                        await _context.IdeaChampion.AddAsync(ideaChampion);
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

        private async Task SendChampionAssignedEmail(ContinuousImprovementChampions champion, ContinuousImprovementIdeas idea)
        {
            var subject = $"Asignación: Nueva Idea de Mejora #{idea.Id}";

            var body = $@"
                <div style='font-family: Arial, sans-serif; color: #333;'>
                    <table width='100%' cellpadding='0' cellspacing='0' style='max-width: 600px; margin: 0 auto; border: 1px solid #e0e0e0; border-radius: 5px;'>
                        <tr>
                            <td style='background-color: #0078D4; padding: 20px; text-align: center; color: white; border-top-left-radius: 5px; border-top-right-radius: 5px;'>
                                <h2 style='margin: 0;'>Nueva Asignación de Idea</h2>
                            </td>
                        </tr>
                        <tr>
                            <td style='padding: 20px;'>
                                <p>Hola <strong>{champion.Name}</strong>,</p>
                                <p>Se te ha asignado el seguimiento de una nueva idea de mejora continua. A continuación los detalles:</p>
                                
                                <table width='100%' cellpadding='10' cellspacing='0' style='background-color: #f9f9f9; border-radius: 5px; margin-top: 10px;'>
                                    <tr>
                                        <td width='30%' style='font-weight: bold; border-bottom: 1px solid #ddd;'>ID:</td>
                                        <td style='border-bottom: 1px solid #ddd;'>#{idea.Id}</td>
                                    </tr>
                                    <tr>
                                        <td style='font-weight: bold; border-bottom: 1px solid #ddd;'>Colaborador:</td>
                                        <td style='border-bottom: 1px solid #ddd;'>{idea.FullName}</td>
                                    </tr>
                                    <tr>
                                        <td style='font-weight: bold; border-bottom: 1px solid #ddd;'>Situación Actual:</td>
                                        <td style='border-bottom: 1px solid #ddd;'>{idea.CurrentSituation}</td>
                                    </tr>
                                    <tr>
                                        <td style='font-weight: bold;'>Propuesta:</td>
                                        <td>{idea.IdeaDescription}</td>
                                    </tr>
                                </table>

                                <p style='margin-top: 20px;'>Por favor, ingresa al sistema para dar seguimiento.</p>
                            </td>
                        </tr>
                        <tr>
                            <td style='background-color: #f0f0f0; padding: 15px; text-align: center; font-size: 12px; color: #666; border-bottom-left-radius: 5px; border-bottom-right-radius: 5px;'>
                                &copy; {DateTime.Now.Year} Equipo de Mejora Continua
                            </td>
                        </tr>
                    </table>
                </div>";

            await _emailService.SendEmailAsync(champion.Email, subject, body);
        }

        private async Task SendNewIdeaRegisteredEmail(ContinuousImprovementIdeas idea)
        {
            var subject = $"Nueva Idea Registrada en el Sistema - Folio #{idea.Id}";

            var body = $@"
                <div style='font-family: Arial, sans-serif; color: #333;'>
                    <table width='100%' cellpadding='0' cellspacing='0' style='max-width: 600px; margin: 0 auto; border: 1px solid #e0e0e0; border-radius: 5px;'>
                        <tr>
                            <td style='background-color: #28a745; padding: 20px; text-align: center; color: white; border-top-left-radius: 5px; border-top-right-radius: 5px;'>
                                <h2 style='margin: 0;'>¡Nueva Idea Registrada!</h2>
                            </td>
                        </tr>
                        <tr>
                            <td style='padding: 20px;'>
                                <p>El sistema ha recibido una nueva propuesta de mejora.</p>
                                
                                <table width='100%' cellpadding='10' cellspacing='0' style='background-color: #f9f9f9; border-radius: 5px; margin-top: 10px;'>
                                    <tr>
                                        <td width='30%' style='font-weight: bold; border-bottom: 1px solid #ddd;'>Folio:</td>
                                        <td style='border-bottom: 1px solid #ddd;'>#{idea.Id}</td>
                                    </tr>
                                    <tr>
                                        <td style='font-weight: bold; border-bottom: 1px solid #ddd;'>Área:</td>
                                        <td style='border-bottom: 1px solid #ddd;'>{idea.WorkArea}</td>
                                    </tr>
                                    <tr>
                                        <td style='font-weight: bold; border-bottom: 1px solid #ddd;'>Registrado por:</td>
                                        <td style='border-bottom: 1px solid #ddd;'>{idea.FullName}</td>
                                    </tr>
                                    <tr>
                                        <td style='font-weight: bold;'>Descripción breve:</td>
                                        <td>{idea.IdeaDescription}</td>
                                    </tr>
                                </table>

                                <p style='margin-top: 20px;'>Es necesario asignar Champions y revisar la viabilidad de la idea.</p>
                            </td>
                        </tr>
                        <tr>
                            <td style='background-color: #f0f0f0; padding: 15px; text-align: center; font-size: 12px; color: #666; border-bottom-left-radius: 5px; border-bottom-right-radius: 5px;'>
                                Notificación automática del Sistema de Mejora Continua
                            </td>
                        </tr>
                    </table>
                </div>";
           
            await _emailService.SendEmailAsync(AdminEmail, subject, body);
        }

    }
}