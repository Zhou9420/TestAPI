using System;
using System.Collections.Generic;
using Nancy.Json;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using Newtonsoft.Json;
using TestAPI.Models;
using Newtonsoft.Json.Linq;

using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.POIFS.FileSystem;
using NPOI.HPSF;
using NPOI.HSSF.Util;
using NPOI.XSSF.UserModel;
using System.IO;
using Nancy;

namespace TestAPI.Controllers
{
    [Route("api/[controller]/[action]")]
    [ApiController]
    public class UsersController : ControllerBase
    {
        private readonly ProductContext _context;                   //数据库上下文

        public UsersController(ProductContext context)
        {
            _context = context;
        }

        //查询
        // GET: api/Users
        [Produces("application/json")]
        [HttpGet]
        public async Task<ActionResult<IEnumerable<User>>> GetUsers(string name)
        {
            if (name==null)
            {
                return await _context.Users.ToListAsync();
            }
            else
            {
                return await _context.Users.Where(p =>p.Name.Contains(name)).ToListAsync();
            }
        }

        //查询单个
        // GET: api/Users/5
        [HttpGet("{id}")]
        public async Task<ActionResult<User>> GetUser(int id)
        {
            var user = await _context.Users.FindAsync(id);

            if (user == null)
            {
                return NotFound();
            }

            return user;
        }

        #region
        /// <summary>
        /// 将object对象转换为实体对象
        /// </summary>
        /// <typeparam name="T">实体对象类名</typeparam>
        /// <param name="asObject">object对象</param>
        /// <returns></returns>
        //public static T ConvertObjectByJson<T>(object asObject) where T : new()
        //{
        //    var serializer = new JavaScriptSerializer();
        //    //将object对象转换为json字符
        //    var json = serializer.Serialize(asObject);
        //    //将json字符转换为实体对象
        //    var t = serializer.Deserialize<T>(json);
        //    return t;
        //}
        #endregion

        //修改
        // PUT: api/Users/5
        // To protect from overposting attacks, enable the specific properties you want to bind to, for
        // more details, see https://go.microsoft.com/fwlink/?linkid=2123754.
        [HttpPut("{id}")]
        public async Task<IActionResult> PutUser(int id,[FromBody]User user)
        {
            if (id != user.Id)
            {
                return BadRequest();
            }

            _context.Entry(user).State = EntityState.Modified;

            try
            {
                await _context.SaveChangesAsync();
            }
            catch (DbUpdateConcurrencyException)
            {
                if (!UserExists(id))
                {
                    return NotFound();
                }
                else
                {
                    throw;
                }
            }

            return NoContent();
        }

        //添加
        // POST: api/Users
        // To protect from overposting attacks, enable the specific properties you want to bind to, for
        // more details, see https://go.microsoft.com/fwlink/?linkid=2123754.
        [HttpPost]
        public async Task<ActionResult<User>> PostUser(User user)
        {
            _context.Users.Add(user);
            await _context.SaveChangesAsync();

            return CreatedAtAction("GetUser", new { id = user.Id }, user);
        }

        //删除
        // DELETE: api/Users/5
        [HttpDelete("{id}")]
        public async Task<ActionResult<User>> DeleteUser(int id)
        {
            var user = await _context.Users.FindAsync(id);
            if (user == null)
            {
                return NotFound();
            }

            _context.Users.Remove(user);
            await _context.SaveChangesAsync();

            return user;
        }

        //批量删除
        [HttpDelete]
        public async Task<ActionResult<int[]>> DeleteAllUser([FromBody] int[] checkedIDS)
        {
            foreach (var item in checkedIDS)
            {
                var user = await _context.Users.FindAsync(item);
                if (user != null)
                {
                    _context.Users.Remove(user);
                }
            }
            await _context.SaveChangesAsync();
            return checkedIDS;
        }

        //导入
        [HttpPost]
        public async Task<string> ExcelIn(IFormFile file)
        {
            string ReturnValue = string.Empty;
            //定义一个bool类型的变量用来做验证
            bool flag = true;
            try
            {
                //获取文件的后缀名，并转换为小写
                string fileExt = Path.GetExtension(file.FileName).ToLower();
                //定义一个集合一会儿将数据存储进来,全部一次丢到数据库中保存
                var Data = new List<User>();
                //将获取到的文件转成为文件流
                MemoryStream ms = new MemoryStream();
                file.CopyTo(ms);
                ms.Seek(0, SeekOrigin.Begin);
                //创建一个 Workbook 对象流(其实就是创建一个Excel文件)
                IWorkbook book;
                //判断获取文件的后缀名
                if (fileExt == ".xlsx")             
                    book = new XSSFWorkbook(ms);
                else if (fileExt == ".xls")
                    book = new HSSFWorkbook(ms);
                else
                    book = null;
                //获取到excel的第一个sheet页
                ISheet sheet = book.GetSheetAt(0);
                //int a = book.NumberOfSheets;获取sheet的数量
                //获取总行数
                int CountRow = sheet.LastRowNum + 1;    //(默认应该从0开始，所以要+1)
                if (CountRow - 1 == 0)
                    return "Excel列表数据项为空!";

                #region 循环验证
                for (int i = 1; i < CountRow; i++)     //循环总行数，从第二行开始(去掉表头)
                {
                    //获取第i行的数据
                    var row = sheet.GetRow(i);
                    if (row != null)
                    {
                        var cell = row.LastCellNum;     //获取每行的单元格个数(也是从0开始)
                        for (int j = 0; j < cell; j++)  //循环的验证单元格中的数据
                        {
                            if (row.GetCell(j) == null || row.GetCell(j).ToString().Trim().Length == 0)
                            {
                                flag = false;
                                ReturnValue += $"第{i + 1}行,第{j + 1}列数据不能为空。";
                            }
                        }
                    }
                }
                #endregion

                if (flag)
                {
                    for (int i = 1; i < CountRow; i++)     //循环获取的总行数
                    {
                        User user = new User();            //实例化实体对象
                        var row = sheet.GetRow(i);         //获取第i行的数据

                        if (row.GetCell(0) != null && row.GetCell(0).ToString().Trim().Length > 0) //如果第i行第一个单元格数据不为空且长度大于0
                        {
                            user.Name = row.GetCell(0).ToString();
                        }
                        if (row.GetCell(1) != null && row.GetCell(1).ToString().Trim().Length > 0)
                        {
                            if (row.GetCell(1).ToString()=="男")
                            {
                                user.Grader = "true";
                            }
                            else
                            {
                                user.Grader = "false";
                            }
                        }
                        if (row.GetCell(2) != null && row.GetCell(2).ToString().Trim().Length > 0)
                        {
                            user.Phone = row.GetCell(2).ToString();
                        }
                        if (row.GetCell(3) != null && row.GetCell(3).ToString().Trim().Length > 0)
                        {
                            user.Address = row.GetCell(3).ToString().ToString();
                        }
                        if (row.GetCell(4) != null && row.GetCell(4).ToString().Trim().Length > 0)
                        {
                            user.Email = row.GetCell(4).ToString();
                        }
                        Data.Add(user);
                    }
                    _context.Users.AddRange(Data);
                    await _context.SaveChangesAsync();

                    ReturnValue = $"数据导入成功,共导入{CountRow - 1}条数据。";
                }

                if (!flag)
                {
                    ReturnValue = "数据存在问题！" + ReturnValue;
                }
            }
            catch (Exception)
            {
                return "服务器异常";
            }
            
            return ReturnValue;
        }





        //导出
        [HttpPost]
        public ActionResult ExcelOut() 
        {
            var users = _context.Users.ToList();

            ////创建工作簿Excel
            HSSFWorkbook book = new HSSFWorkbook();
            //为工作簿创建工作表并命名
            NPOI.SS.UserModel.ISheet sheet1 = book.CreateSheet("用户信息");
            //创建第一行
            NPOI.SS.UserModel.IRow row1 = sheet1.CreateRow(0);
            //创建其他列并赋值(相当于表格的表头)
            row1.CreateCell(0).SetCellValue("姓名");
            row1.CreateCell(1).SetCellValue("性别");
            row1.CreateCell(2).SetCellValue("班级");
            row1.CreateCell(3).SetCellValue("地址");
            row1.CreateCell(4).SetCellValue("邮箱");
            //循环输出
            for (int i = 0; i < users.Count(); i++)
            {
                //创建行
                NPOI.SS.UserModel.IRow rowTemp = sheet1.CreateRow(i + 1);
                rowTemp.CreateCell(0).SetCellValue(users[i].Name);
                if (users[i].Grader=="true")
                    rowTemp.CreateCell(1).SetCellValue("男");
                else
                    rowTemp.CreateCell(1).SetCellValue("女");
                rowTemp.CreateCell(2).SetCellValue(users[i].Phone);
                rowTemp.CreateCell(3).SetCellValue(users[i].Address);
                rowTemp.CreateCell(4).SetCellValue(users[i].Email);
            }
            //文件名
            var fileName = "用户信息报表" + DateTime.Now.ToString("yyyy-MM-dd") + ".xls";
            //将Excel表格转化为流，输出
            MemoryStream bookStream = new MemoryStream();//创建文件流
            book.Write(bookStream); //文件写入流（向流中写入字节序列）
            bookStream.Seek(0, SeekOrigin.Begin);//输出之前调用Seek，把0位置指定为开始位置
            return File(bookStream, "application/vnd.ms-excel", fileName);//最后以文件形式返回
        }


        //判断是否存在
        private bool UserExists(int id)
        {
            return _context.Users.Any(e => e.Id == id);
        }


    }
}
