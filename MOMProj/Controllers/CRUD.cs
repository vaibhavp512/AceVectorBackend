using Call.Modles.Entities;
using Microsoft.AspNetCore.Mvc;
using System.Collections.Generic;

namespace YourNamespace.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class CRUD : ControllerBase
    {
        [HttpGet("model1/{id}")]
        public ActionResult<List<MeetingMOMDetails>> GetModel1(int id)
        {
            var excelService = new ExcelService<MeetingMOMDetails>("C:\\Users\\Public\\Excel\\MeetingMOMDetails.xlsx");
            if(excelService.ReadExcelDataMom(id).Count== 0)
            {
                return NotFound("Not data found");
            }
            else
            return Ok(excelService.ReadExcelDataMom(id));
        }

        [HttpGet("model2/{subject}")]
        public ActionResult<List<MeetingMaster>> GetModel2(string subject)
        {
            var excelService = new ExcelService<MeetingMaster>("C:\\Users\\Public\\Excel\\MeetingMaster.xlsx");
            return Ok(excelService.ReadExcelDataSubject(subject));
        }

     
        //This is for getting meeting participant for sending email will pass id to search
        [HttpGet("model3")]
        public ActionResult<List<MeetingAttendes>> GetModel3()
        {
            var excelService = new ExcelService<MeetingAttendes>("C:\\Users\\Public\\Excel\\MeetingAttendes.xlsx");
            return Ok(excelService.ReadExcelData());
        }

        //Below  method will be used to save or upadet MOM detail, we will parameter like Mid	MOM	Updatedby	Timestamp
        //for first time will check we will add directly whiout checking 
        //For next save we will check the current MOM with the last avilable entry in DB based on Timestamp 
        //If different we will create a new entry in DB.
        [HttpPost("model1")] 
        public IActionResult AddModel1([FromBody] MeetingMOMDetails model)
        {
            var excelService = new ExcelService<MeetingMOMDetails>("C:\\Users\\Public\\Excel\\MeetingMOMDetails.xlsx");
            
//            excelService.AddDataVal(model);
            //excelService.AddDataValidation(model);
            excelService.AddData(model);
            return Ok();
        }

        //[HttpPost("model2")]
        //public IActionResult AddModel2([FromBody] MeetingMaster model)
        //{
        //    var excelService = new ExcelService<MeetingMaster>("C:\\Users\\Public\\Excel\\MeetingMaster.xlsx");
        //    excelService.AddData(model);
        //    return Ok();
        //}
        //   [HttpPost("model3")]
        //public IActionResult AddModel3([FromBody] MeetingAttendes model)
        //{
        //    var excelService = new ExcelService<MeetingAttendes>("C:\\Users\\Public\\Excel\\MeetingAttendes.xlsx");
        //    excelService.AddData(model);
        //    return Ok();
        //}


        //Below method will be used to update method mom detail will pass same parameter as post method 
        [HttpPut("model1/{id}")]
        public IActionResult UpdateModel1(int id, [FromBody] MeetingMOMDetails model)
        {
            var excelService = new ExcelService<MeetingMOMDetails>("C:\\Users\\Public\\Excel\\MeetingMOMDetails.xlsx");
            excelService.UpdateData(model, id);
            return Ok();
        }

        //[HttpPut("model2/{id}")]
        //public IActionResult UpdateModel2(int id, [FromBody] MeetingMaster model)
        //{
        //    var excelService = new ExcelService<MeetingMaster>("C:\\Users\\Public\\Excel\\MeetingMaster.xlsx");
        //    excelService.UpdateData(model, id);
        //    return Ok();
        //} 
        //[HttpPut("model3/{id}")]
        //public IActionResult UpdateModel3(int id, [FromBody] MeetingAttendes model)
        //{
        //    var excelService = new ExcelService<MeetingAttendes>("C:\\Users\\Public\\Excel\\MeetingAttendes.xlsx");
        //    excelService.UpdateData(model, id);
        //    return Ok();
        //}

        //[HttpDelete("model1/{id}")]
        //public IActionResult DeleteModel1(int id)
        //{
        //    var excelService = new ExcelService<MeetingMOMDetails>("C:\\Users\\Public\\Excel\\MeetingMOMDetails.xlsx");
        //    excelService.DeleteData(id);
        //    return Ok();
        //}

        //[HttpDelete("model2/{id}")]
        //public IActionResult DeleteModel2(int id)
        //{
        //    var excelService = new ExcelService<MeetingMaster>("C:\\Users\\Public\\Excel\\MeetingMaster.xlsx");
        //    excelService.DeleteData(id);
        //    return Ok();
        //}
        //[HttpDelete("model3/{id}")]
        //public IActionResult DeleteModel3(int id)
        //{
        //    var excelService = new ExcelService<MeetingAttendes>("C:\\Users\\Public\\Excel\\MeetingAttendes.xlsx");
        //    excelService.DeleteData(id);
        //    return Ok();
        //}
    }
}
