using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Microsoft.AspNetCore.Mvc;
using System.Linq;

namespace GoogleSheetsAPI.Controllers
{
[Route("api/[controller]")]
[ApiController]
public class ItemsController : ControllerBase
{
    const string SPREADSHEET_ID = "143JhAuG7l....";
    const string SHEET_NAME = "Items";

    GoogleSheetsHelper _googleSheetsHelper;

    public ItemsController(GoogleSheetsHelper googleSheetsHelper)
    {
        _googleSheetsHelper = googleSheetsHelper ?? throw new System.ArgumentNullException(nameof(googleSheetsHelper));
    }

    // GET: api/<ItemsController>
    [HttpGet]
    public IActionResult Get()
    {
        var range = $"{SHEET_NAME}!A:D";
        var request = _googleSheetsHelper.Service.Spreadsheets.Values.Get(SPREADSHEET_ID, range);
        var response = request.Execute();
        var values = response.Values;

        return Ok(ItemsMapper.MapFromRangeData(values));
    }

    // GET api/<ItemsController>/5
    [HttpGet("{rowId}")]
    public IActionResult Get(int rowId)
    {
        var range = $"{SHEET_NAME}!A{rowId}:D{rowId}";
        var request = _googleSheetsHelper.Service.Spreadsheets.Values.Get(SPREADSHEET_ID, range);
        var response = request.Execute();
        var values = response.Values;

        return Ok(ItemsMapper.MapFromRangeData(values).FirstOrDefault());
    }

    // POST api/<ItemsController>
    [HttpPost]
    public void Post(Item item)
    {
        var range = $"{SHEET_NAME}!A:D";
        var valueRange = new ValueRange
        {
            Values = ItemsMapper.MapToRangeData(item)
        };

        var appendRequest = _googleSheetsHelper.Service.Spreadsheets.Values.Append(valueRange, SPREADSHEET_ID, range);
        appendRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
        appendRequest.Execute();            
    }

    // PUT api/<ItemsController>/5
    [HttpPut("{rowId}")]
    public void Put(int rowId, Item item)
    {
        var range = $"{SHEET_NAME}!A{rowId}:D{rowId}";
        var valueRange = new ValueRange
        {
            Values = ItemsMapper.MapToRangeData(item)
        };

        var updateRequest = _googleSheetsHelper.Service.Spreadsheets.Values.Update(valueRange, SPREADSHEET_ID, range);
        updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
        updateRequest.Execute();
    }

    // DELETE api/<ItemsController>/5
    [HttpDelete("{rowId}")]
    public void Delete(int rowId)
    {
        var range = $"{SHEET_NAME}!A{rowId}:D{rowId}";
        var requestBody = new ClearValuesRequest();

        var deleteRequest = _googleSheetsHelper.Service.Spreadsheets.Values.Clear(requestBody, SPREADSHEET_ID, range);
        deleteRequest.Execute();
    }
}
}
