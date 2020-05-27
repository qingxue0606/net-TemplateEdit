using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;

namespace TemplateEdit.Controllers
{
    public class WordController : Controller
    {
        private readonly IWebHostEnvironment _webHostEnvironment;

        public WordController(IWebHostEnvironment webHostEnvironment)
        {
            _webHostEnvironment = webHostEnvironment;
        }

        public IActionResult FileMaker()
        {
            PageOfficeNetCore.WordWriter.WordDocument doc = new PageOfficeNetCore.WordWriter.WordDocument();
            doc.OpenDataRegion("Name").Value = "王五";
            doc.OpenDataRegion("Name").Editing = true;// docSubmitForm提交模式打开文件的话，此区域可以编辑
            doc.OpenDataRegion("Address").Value = "上海市xx区南xxx路xxx号";
            doc.OpenDataRegion("Tel").Value = "021-66662222";
            doc.OpenDataRegion("Phone").Value = "13811112222";
            doc.OpenDataRegion("Sex").Value = "男";
            doc.OpenDataRegion("Age").Value = "28";
            doc.OpenDataTag("{ 甲方公司名称 }").Value = "北京联想公司";
            doc.OpenDataTag("{ 乙方公司名称 }").Value = "北京幻想科技公司";
            doc.OpenDataTag("【 合同日期 】").Value = "2014年08月01日";
            doc.OpenDataTag("【 合同编号 】").Value = "201408010001";
            PageOfficeNetCore.FileMakerCtrl fileMakerCtrl = new PageOfficeNetCore.FileMakerCtrl(Request);
            fileMakerCtrl.ServerPage = "/PageOffice/POServer";

            fileMakerCtrl.SaveFilePage = "SaveDoc?type=2";
            fileMakerCtrl.SetWriter(doc);
            fileMakerCtrl.JsFunction_OnProgressComplete = "OnProgressComplete()";
            fileMakerCtrl.FillDocument("/doc/test.doc", PageOfficeNetCore.DocumentOpenType.Word);

            ViewBag.FMCtrl = fileMakerCtrl.GetHtmlCode("FileMakerCtrl1");
            return View();
        }

        public IActionResult openMaker()
        {
            string webRootPath = _webHostEnvironment.WebRootPath;
            ViewBag.url = _webHostEnvironment.WebRootPath + "\\doc\\filemaker.doc";
            return View();
        }

        public IActionResult template()
        {
            PageOfficeNetCore.PageOfficeCtrl pageofficeCtrl = new PageOfficeNetCore.PageOfficeCtrl(Request);
            pageofficeCtrl.ServerPage = "../PageOffice/POServer";
            PageOfficeNetCore.WordWriter.WordDocument doc = new PageOfficeNetCore.WordWriter.WordDocument();
            doc.Template.DefineDataRegion("Name", "[ 担保人姓名 ]");
            doc.Template.DefineDataRegion("Address", "[ 担保人地址 ]");
            doc.Template.DefineDataRegion("Tel", "[ 担保人电话 ]");
            doc.Template.DefineDataRegion("Phone", "[ 担保人手机 ]");
            doc.Template.DefineDataRegion("Sex", "[ 担保人性别 ]");
            doc.Template.DefineDataRegion("Age", "[ 担保人年龄 ]");
            doc.Template.DefineDataTag("{ 甲方公司名称 }");
            doc.Template.DefineDataTag("{ 乙方公司名称 }");
            doc.Template.DefineDataTag("【 合同日期 】");
            doc.Template.DefineDataTag("【 合同编号 】");
            pageofficeCtrl.AddCustomToolButton("保存", "Save()", 1);
            pageofficeCtrl.AddCustomToolButton("定义数据区域", "ShowDefineDataRegions()", 3);
            pageofficeCtrl.AddCustomToolButton("定义数据标签", "ShowDefineDataTags()", 20);
            pageofficeCtrl.Theme = PageOfficeNetCore.ThemeType.Office2007;
            pageofficeCtrl.BorderStyle = PageOfficeNetCore.BorderStyleType.BorderThin;
            pageofficeCtrl.SetWriter(doc);
            //设置保存页面
            pageofficeCtrl.SaveFilePage = "SaveDoc";
            //打开Word文档
            pageofficeCtrl.WebOpen("/doc/test.doc", PageOfficeNetCore.OpenModeType.docNormalEdit, "tom");
            ViewBag.POCtrl = pageofficeCtrl.GetHtmlCode("PageOfficeCtrl1");

            return View();
        }
        public IActionResult open()
        {
            PageOfficeNetCore.PageOfficeCtrl pageofficeCtrl = new PageOfficeNetCore.PageOfficeCtrl(Request);
            pageofficeCtrl.ServerPage = "../PageOffice/POServer";
            string type = Request.Query["type"].ToString();

            PageOfficeNetCore.WordWriter.WordDocument doc = new PageOfficeNetCore.WordWriter.WordDocument();
            doc.EnableAllDataRegionsEditing = true; // 此属性可以设置在提交模式（docSubmitForm）下，所有的数据区域可以编辑

            doc.OpenDataRegion("Name").Value = "张三";
            doc.OpenDataRegion("Address").Value = "北京市丰台区南四环西路xxx号";
            doc.OpenDataRegion("Tel").Value = "010-88882222";
            doc.OpenDataRegion("Phone").Value = "13822225555";
            doc.OpenDataRegion("Sex").Value = "男";
            doc.OpenDataRegion("Age").Value = "21";
            doc.OpenDataTag("{ 甲方公司名称 }").Value = "微软中国总部";
            doc.OpenDataTag("{ 乙方公司名称 }").Value = "北京幻想科技公司";
            doc.OpenDataTag("【 合同日期 】").Value = "2014年08月01日";
            doc.OpenDataTag("【 合同编号 】").Value = "201408010001";

            pageofficeCtrl.SetWriter(doc);

            if ("1" == type)
            {
                pageofficeCtrl.AddCustomToolButton("保存", "Save2()", 1);
                pageofficeCtrl.WebOpen("/doc/test.doc", PageOfficeNetCore.OpenModeType.docNormalEdit, "zhangsan");
            }
            else if ("2" == type)
            {
                pageofficeCtrl.Menubar = true;
                pageofficeCtrl.CustomToolbar = false;
                pageofficeCtrl.OfficeToolbars = false;
                pageofficeCtrl.WebOpen("/doc/test.doc", PageOfficeNetCore.OpenModeType.docReadOnly, "zhangsan");
            }
            else if ("3" == type)
            {
                pageofficeCtrl.AddCustomToolButton("保存", "Save()", 1);
                pageofficeCtrl.SaveDataPage = "SaveData";
                pageofficeCtrl.OfficeToolbars = false;
                pageofficeCtrl.WebOpen("/doc/test.doc", PageOfficeNetCore.OpenModeType.docSubmitForm, "zhangsan");
            }
            ViewBag.POCtrl = pageofficeCtrl.GetHtmlCode("PageOfficeCtrl1");

            return View();
        }

        public async Task<ActionResult> SaveData()
        {
            string content = "";
            PageOfficeNetCore.WordReader.WordDocument doc = new PageOfficeNetCore.WordReader.WordDocument(Request, Response);
            await doc.LoadAsync();
            //获取提交的数值
            PageOfficeNetCore.WordReader.DataRegion poName = doc.OpenDataRegion("PO_Name");
            try
            {
                PageOfficeNetCore.WordReader.DataRegion dataDeptName = doc.OpenDataRegion("PO_deptName");
                content += "后台获取 PO_Name的值：" + poName.Value;
            }
            catch
            {
                content += "客户端提交的数据区域中没有包含名称为 PO_Name 的数据区域。";
            }
            await Response.Body.WriteAsync(Encoding.GetEncoding("GB2312").GetBytes(content));

            doc.ShowPage(400, 300);
            doc.Close();
            return Content("OK");
        }
        public async Task<ActionResult> SaveDoc()
        {


            string type = Request.Query["type"].ToString();
            PageOfficeNetCore.FileSaver fs = new PageOfficeNetCore.FileSaver(Request, Response);
            await fs.LoadAsync();
            string webRootPath = _webHostEnvironment.WebRootPath;
            if ("2" == type)
            {
                fs.SaveToFile(webRootPath + "/doc/" + "filemaker.doc");
            }
            else 
            {
                fs.SaveToFile(webRootPath + "/doc/" + fs.FileName);
            }
            fs.Close();
            return Content("OK");
        }

    }
}