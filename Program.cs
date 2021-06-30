using System;
using System.Collections.Generic;

namespace ReportPublish
{
    class Program
    {
        static void Main(string[] args)
        {

            GeneratedClass docx = new GeneratedClass();
            GiaoVien gv = new GiaoVien("Nguyen Van A", "GV001");
            ThongTinGiangDay[] t = new ThongTinGiangDay[30];
            for (int i = 0; i < 30; ++i)
            {
                t[i].MonHoc = "mon" + i.ToString();
                t[i].BacDaoTao = "DH";
                t[i].Lop = "lop " + i.ToString();
                t[i].GioChuan = 2.5;
            }

            for (int i = 0; i < 30; ++i)
            {
                gv.Ttgd.Add(t[i]);
            }
            docx.CreatePackage("E:\\demo.docx", gv);
        }


    }
}
