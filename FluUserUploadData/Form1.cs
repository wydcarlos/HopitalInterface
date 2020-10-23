using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.OracleClient;
using System.IO;
using ICSharpCode.SharpZipLib.Zip;
using ICSharpCode.SharpZipLib.Checksums;
using System.Threading;
using System.Xml.Linq;

namespace FluUserUploadData
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        public static string ConnectionString = "data source=bjhis;password=his-0765;persist security info=True;user id=bjhis";
        public string strZip = "";
        public string strFile = AppDomain.CurrentDomain.BaseDirectory + @"\流感系统\Excel";

        public DataSet FluDataSet = new DataSet();

        /// <summary>
        /// 门急诊和在院流感病例
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button1_Click(object sender, EventArgs e)
        {

            #region 门急诊和在院流感病例数据
            string sql = @"


SELECT '45608783744060611A1001' P900, -- 医疗机构代码  是
       '佛山市顺德区北滘医院' P6891, -- 机构名称  是
       NVL(A.MCARD_NO, '-') P686, -- 医疗保险手册（卡）号  
       '' P800, --健康卡号 
       '01' P7501, --就诊类型   
       A.CARD_NO P7502, -- 就诊卡号     门急诊卡号或住院号
       A.NAME P4, --姓名    
       DECODE(A.SEX_CODE, 'F', '2', 'M', '1', '9') P5, --性别    
       (CASE
         WHEN A.BIRTHDAY < TO_DATE('1900-01-01', 'yyyy-mm-dd') THEN
          TO_DATE('1900-01-01', 'yyyy-mm-dd')
         ELSE
          A.BIRTHDAY
       END) P6, --出生日期,
       (CASE
         WHEN (CASE
                WHEN A.BIRTHDAY < TO_DATE('1900-01-01', 'yyyy-mm-dd') THEN
                 0
                ELSE
                 (CASE
                   WHEN ADD_MONTHS(A.BIRTHDAY,
                                   (EXTRACT(YEAR FROM A.REG_DATE) -
                                   EXTRACT(YEAR FROM A.BIRTHDAY)) * 12) < A.REG_DATE THEN
                    EXTRACT(YEAR FROM A.REG_DATE) - EXTRACT(YEAR FROM A.BIRTHDAY)
                   ELSE
                    EXTRACT(YEAR FROM A.REG_DATE) - EXTRACT(YEAR FROM A.BIRTHDAY) - 1
                 END)
              END) < 0 THEN
          0
         ELSE
          (CASE
            WHEN A.BIRTHDAY < TO_DATE('1900-01-01', 'yyyy-mm-dd') THEN
             0
            ELSE
             (CASE
               WHEN ADD_MONTHS(A.BIRTHDAY,
                               (EXTRACT(YEAR FROM A.REG_DATE) -
                               EXTRACT(YEAR FROM A.BIRTHDAY)) * 12) < A.REG_DATE THEN
                EXTRACT(YEAR FROM A.REG_DATE) - EXTRACT(YEAR FROM A.BIRTHDAY)
               ELSE
                EXTRACT(YEAR FROM A.REG_DATE) - EXTRACT(YEAR FROM A.BIRTHDAY) - 1
             END)
          END)
       END) P7, --年龄    
       '01' P7503, --注册证件类型代码  
       NVL(A.IDENNO, '-') P13, --注册证件号码 
       (SELECT O.INTERFACE_CODE
          FROM COM_COMPARE_INTERFACEINFO O
         WHERE A.DEPT_CODE = O.LOCAL_CODE
           AND O.INTERFACE_TYPE = '科室对照'
           AND ROWNUM = 1) P7504, --13  就诊科室代码    
       (SELECT COUNT(*) + 1
          FROM FIN_OPR_REGISTER FOOR
         WHERE FOOR.CARD_NO = A.CARD_NO
           AND FOOR.REG_DATE < A.REG_DATE
           AND A.TRANS_TYPE = '1'
           AND A.VALID_FLAG = '1') P7505, --就诊次数
       A.REG_DATE P7506, -- 15 就诊日期    
       NVL((SELECT ITEM.DIAG_NAME
             FROM MET_CAS_DIAGNOSE ITEM
            WHERE ITEM.INPATIENT_NO = A.CLINIC_CODE
              AND ((INSTR(ITEM.DIAG_NAME, '甲') > 0 AND
                  INSTR(ITEM.DIAG_NAME, '流') > INSTR(ITEM.DIAG_NAME, '甲') AND
                  ITEM.DIAG_NAME NOT LIKE '%流感嗜血%' AND
                  ITEM.DIAG_NAME NOT LIKE '%副流感%' AND
                  ITEM.DIAG_NAME NOT LIKE '%血流感染%') OR
                  (INSTR(ITEM.DIAG_NAME, '乙') > 0 AND
                  INSTR(ITEM.DIAG_NAME, '流') > INSTR(ITEM.DIAG_NAME, '乙') AND
                  ITEM.DIAG_NAME NOT LIKE '%流感嗜血%' AND
                  ITEM.DIAG_NAME NOT LIKE '%副流感%' AND
                  ITEM.DIAG_NAME NOT LIKE '%血流感染%') OR
                  (INSTR(ITEM.DIAG_NAME, '流') > 0 AND
                  INSTR(ITEM.DIAG_NAME, '感') > INSTR(ITEM.DIAG_NAME, '流') AND
                  ITEM.DIAG_NAME NOT LIKE '%流感嗜血%' AND
                  ITEM.DIAG_NAME NOT LIKE '%副流感%' AND
                  ITEM.DIAG_NAME NOT LIKE '%血流感染%') OR
                  (INSTR(ITEM.DIAG_NAME, 'H') > 0 AND
                  INSTR(ITEM.DIAG_NAME, 'N') > INSTR(ITEM.DIAG_NAME, 'H')) OR
                  ((INSTR(ITEM.DIAG_NAME, '高热') > 0 OR
                  INSTR(ITEM.DIAG_NAME, '发烧') > 0 OR
                  INSTR(ITEM.DIAG_NAME, '高热') > 0) AND
                  (INSTR(ITEM.DIAG_NAME, '咳嗽') > 0 OR
                  INSTR(ITEM.DIAG_NAME, '咳痰') > 0)) OR
                  (ITEM.ICD_CODE IN
                  (SELECT ITEM.ICD_CODE
                       FROM MET_COM_ICD10 A
                      WHERE ITEM.ICD_CODE > 'J00'
                        AND ITEM.ICD_CODE < 'J99')) -- 通过诊断编码
                  )
              AND ROWNUM = 1),
           '-') P7507, --主诉
       (SELECT MCD.ICD_CODE
          FROM MET_CAS_DIAGNOSE MCD
         WHERE MCD.INPATIENT_NO = A.CLINIC_CODE
              -- AND MCD.DIAG_KIND = '7'
           AND ROWNUM = 1) P321, --主要疾病诊断代码 
       NVL((SELECT MCD.DIAG_NAME
             FROM MET_CAS_DIAGNOSE MCD
            WHERE MCD.INPATIENT_NO = A.CLINIC_CODE
                 --AND MCD.DIAG_KIND = '7'
              AND ROWNUM = 1),
           '-') P322, --主要疾病诊断描述
       '' P324, --主要疾病诊断代码1
       '' P325, --主要疾病诊断描述1
       '' P327, --主要疾病诊断代码2
       '' P328, --主要疾病诊断描述2
       '' P3291, --主要疾病诊断代码3
       '' P3292, --主要疾病诊断描述3
       '' P3294, --主要疾病诊断代码4
       '' P3295, --主要疾病诊断描述4
       '' P3297, --主要疾病诊断代码5
       '' P3298, --主要疾病诊断描述5
       '' P3281, --主要疾病诊断代码6
       '' P3282, --主要疾病诊断描述6
       '' P3284, --主要疾病诊断代码7 
       '' P3285, --主要疾病诊断描述7 
       '' P3287, --主要疾病诊断代码8
       '' P3288, --主要疾病诊断描述8
       '' P3271, --主要疾病诊断代码9
       '' P3272, --主要疾病诊断描述9
       '' P3274, --主要疾病诊断代码10
       '' P3275, --主要疾病诊断描述10    
       '' P6911, --重症监护室名称 1
       '' P6912, --进入时间 1   
       '' P6913, --退出时间 1
       '' P6914, --重症监护室名称 2
       '' P6915, --进入时间 2   
       '' P6916, --退出时间 2
       '' P6917, --重症监护室名称 3
       '' P6918, --进入时间 3   
       '' P6919, --退出时间 3
       '' P6920, --重症监护室名称 4
       '' P6921, --进入时间 4  
       '' P6922, --退出时间 4    
       '' P6923, --重症监护室名称 4
       '' P6924, --进入时间 4  
       '' P6925, --退出时间 4    
       DECODE(A.PACT_CODE, '1', '7', '15', '1', '9') P1, --医疗费用支付方式代码 
       NVL((SELECT SUM(FOF.PUB_COST + FOF.PAY_COST + FOF.OWN_COST)
             FROM FIN_OPB_FEEDETAIL FOF
            WHERE FOF.CLINIC_CODE = A.CLINIC_CODE),
           0) P7508, --总费用   条件必填仅门急诊病例填写,单位：人民币元
       A.OWN_COST P7509, --56 挂号费  条件必填仅门急诊病例填写,单位：人民币元
       NVL((SELECT SUM(FOF.PUB_COST + FOF.PAY_COST + FOF.OWN_COST)
             FROM FIN_OPB_FEEDETAIL FOF
            WHERE FOF.CLINIC_CODE = A.CLINIC_CODE
              AND FOF.DRUG_FLAG = '1'),
           0) P7510, -- 药品费 条件必填仅门急诊病例填写,单位：人民币元
       NVL((SELECT SUM(FOF.PUB_COST + FOF.PAY_COST + FOF.OWN_COST)
             FROM FIN_OPB_FEEDETAIL FOF
            WHERE FOF.CLINIC_CODE = A.CLINIC_CODE
              AND FOF.CLASS_CODE = 'UC'),
           0) P7511, --58 检查费  条件必填仅门急诊病例填写,单位：人民币元
       NVL((SELECT SUM(FOF.OWN_COST)
             FROM FIN_OPB_FEEDETAIL FOF
            WHERE FOF.CLINIC_CODE = A.CLINIC_CODE),
           0) P7512, --59 自付费用 条件必填仅门急诊病例填写,单位：人民币元
       '2' P8508, --60 是否死亡  
       '' P8509 --61 死亡时间  
  FROM FIN_OPR_REGISTER A
 WHERE A.VALID_FLAG = '1'
   AND A.CLINIC_CODE = '{0}'

UNION ALL

SELECT '45608783744060611A1001' P900, -- 医疗机构代码  是
       '佛山市顺德区北滘医院' P6891, -- 机构名称  是
       NVL(FII.MCARD_NO, '-') P686, -- 医疗保险手册（卡）号  
       '' P800, --健康卡号 
       '03' P7501, --就诊类型   
       FII.PATIENT_NO P7502, -- 就诊卡号   门急诊卡号或住院号
       FII.NAME P4, --姓名    
       DECODE(FII.SEX_CODE, 'F', '2', 'M', '1', '9') P5, --性别  
       (CASE
         WHEN FII.BIRTHDAY < TO_DATE('1900-01-01', 'yyyy-mm-dd') THEN
          TO_DATE('1900-01-01', 'yyyy-mm-dd')
         ELSE
          FII.BIRTHDAY
       END) P6, --出生日期,
       (CASE
         WHEN (CASE
                WHEN FII.BIRTHDAY < TO_DATE('1900-01-01', 'yyyy-mm-dd') THEN
                 0
                ELSE
                 (CASE
                   WHEN ADD_MONTHS(FII.BIRTHDAY,
                                   (EXTRACT(YEAR FROM FII.IN_DATE) -
                                   EXTRACT(YEAR FROM FII.BIRTHDAY)) * 12) <
                        FII.IN_DATE THEN
                    EXTRACT(YEAR FROM FII.IN_DATE) - EXTRACT(YEAR FROM FII.BIRTHDAY)
                   ELSE
                    EXTRACT(YEAR FROM FII.IN_DATE) - EXTRACT(YEAR FROM FII.BIRTHDAY) - 1
                 END)
              END) < 0 THEN
          0
         ELSE
          (CASE
            WHEN FII.BIRTHDAY < TO_DATE('1900-01-01', 'yyyy-mm-dd') THEN
             0
            ELSE
             (CASE
               WHEN ADD_MONTHS(FII.BIRTHDAY,
                               (EXTRACT(YEAR FROM FII.IN_DATE) -
                               EXTRACT(YEAR FROM FII.BIRTHDAY)) * 12) <
                    FII.IN_DATE THEN
                EXTRACT(YEAR FROM FII.IN_DATE) - EXTRACT(YEAR FROM FII.BIRTHDAY)
               ELSE
                EXTRACT(YEAR FROM FII.IN_DATE) - EXTRACT(YEAR FROM FII.BIRTHDAY) - 1
             END)
          END)
       END) P7, --年龄    
       '01' P7503, --注册证件类型代码  
       NVL(FII.IDENNO, '-') P13, --注册证件号码     
       (SELECT O.INTERFACE_CODE
          FROM COM_COMPARE_INTERFACEINFO O
         WHERE FII.DEPT_CODE = O.LOCAL_CODE
           AND O.INTERFACE_TYPE = '科室对照'
           AND ROWNUM = 1) P7504, --13  就诊科室代码    
       TO_NUMBER(SUBSTR(FII.INPATIENT_NO, 3, 3)) P7505, --就诊次数
       FII.IN_DATE P7506, -- 15 就诊日期    
       '-' P7507, --主诉
       (SELECT MCD.ICD_CODE
          FROM MET_CAS_DIAGNOSE MCD
         WHERE MCD.INPATIENT_NO = FII.INPATIENT_NO
           AND ROWNUM = 1) P321, --主要疾病诊断代码
       NVL((SELECT MCD.DIAG_NAME
             FROM MET_CAS_DIAGNOSE MCD
            WHERE MCD.INPATIENT_NO = FII.INPATIENT_NO
              AND ROWNUM = 1),
           '-') P322, --主要疾病诊断描述
       '' P324, --主要疾病诊断代码1
       '' P325, --主要疾病诊断描述1
       '' P327, --主要疾病诊断代码2
       '' P328, --主要疾病诊断描述2
       '' P3291, --主要疾病诊断代码3
       '' P3292, --主要疾病诊断描述3
       '' P3294, --主要疾病诊断代码4
       '' P3295, --主要疾病诊断描述4
       '' P3297, --主要疾病诊断代码5
       '' P3298, --主要疾病诊断描述5
       '' P3281, --主要疾病诊断代码6
       '' P3282, --主要疾病诊断描述6
       '' P3284, --主要疾病诊断代码7 
       '' P3285, --主要疾病诊断描述7 32
       '' P3287, --主要疾病诊断代码8
       '' P3288, --主要疾病诊断描述8
       '' P3271, --主要疾病诊断代码9
       '' P3272, --主要疾病诊断描述9
       '' P3274, --主要疾病诊断代码10
       '' P3275, --主要疾病诊断描述10    38
       '' P6911, --重症监护室名称 1
       '' P6912, --进入时间 1   
       '' P6913, --退出时间 1
       '' P6914, --重症监护室名称 2
       '' P6915, --进入时间 2   
       '' P6916, --退出时间 2
       '' P6917, --重症监护室名称 3
       '' P6918, --进入时间 3   
       '' P6919, --退出时间 3
       '' P6920, --重症监护室名称 4
       '' P6921, --进入时间 4  
       '' P6922, --退出时间 4    50
       '' P6923, --重症监护室名称 4
       '' P6924, --进入时间 4  
       '' P6925, --退出时间 4    53
       '' P1, --医疗费用支付方式代码 54 条件仅门急诊病例填写
       0 P7508, --55 总费用  条件仅门急诊病例填写
       0 P7509, --56 挂号费  条件必填仅门急诊病例填写
       0 P7510, --57 药品费 条件必填仅门急诊病例填写
       0 P7511, --58 检查费  条件必填仅门急诊病例填写
       0 P7512, --59 自付费用  条件必填仅门急诊病例填写
       '2' P8508, --60 是否死亡  
       '' P8509 --61 死亡时间  
  FROM FIN_IPR_INMAININFO FII
 WHERE FII.INPATIENT_NO = '{0}'
   AND FII.IN_STATE IN ('R', 'I', 'P')


";
            #endregion

            #region 导出数据

            //获取流感患者
            DataSet vfp = FluDataSet;

            DataTable all = new DataTable();

            if (vfp.Tables[0].Rows.Count > 0)
            {

                DataTable dt = vfp.Tables[0];

                var query = from t in dt.AsEnumerable()
                            group t by new { t1 = t.Field<string>("clinic_code") } into m
                            select new
                            {
                                clinic_code = m.Key.t1
                            };
                if (query.ToList().Count > 0)
                {
                    query.ToList().ForEach(q =>
                    {
                        if (!q.clinic_code.Contains("-"))
                        {
                            DataSet ds = GetDataSet(string.Format(sql, q.clinic_code));

                            if (ds.Tables[0].Rows.Count > 0)
                            {
                                all.Merge(ds.Tables[0]);
                            }
                        }
                    });
                }

                if (all.Rows.Count > 0)
                {
                    string filePath = strFile + @"\flu_" + System.DateTime.Now.ToString("yyyyMMddHHmmss") + ".csv";

                    TableToCsv(all, filePath);
                    strZip = filePath.Replace("Excel", "ZipFile");
                    

                    //压缩备份
                    string strZipBak = strZip.Replace("ZipFile", "ZipFileBak").Replace(".csv", ".zip");
                    ZipFile(strFile, strZipBak);

                    File.Copy(filePath, strZip, true);
                }
            }
            #endregion

        }

        /// <summary>
        /// 导出csv文件
        /// </summary>
        /// <param name="dt"></param>
        /// <param name="filePath"></param>
        public static void TableToCsv(DataTable dt, string filePath)
        {
            FileInfo fi = new FileInfo(filePath);
            string path = fi.DirectoryName;
            string name = fi.Name;
            //\/:*?"<>|
            //把文件名和路径分别取出来处理
            name = name.Replace(@"\", "＼");
            name = name.Replace(@"/", "／");
            name = name.Replace(@":", "：");
            name = name.Replace(@"*", "＊");
            name = name.Replace(@"?", "？");
            name = name.Replace(@"<", "＜");
            name = name.Replace(@">", "＞");
            name = name.Replace(@"|", "｜");
            string title = "";

            FileStream fs = new FileStream(path + "\\" + name, FileMode.Create);
            StreamWriter sw = new StreamWriter(new BufferedStream(fs), System.Text.Encoding.Default);

            for (int i = 0; i < dt.Columns.Count; i++)
            {
                title += dt.Columns[i].ColumnName + ",";
            }
            title = title.Substring(0, title.Length - 1) + "\n";
            sw.Write(title);

            foreach (DataRow row in dt.Rows)
            {
                if (row.RowState == DataRowState.Deleted) continue;
                string line = "";
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    line += row[i].ToString().Replace(",", "") + ",\t";//加\t为设置单元格文本格式
                }
                line = line.Substring(0, line.Length - 1) + "\n";

                sw.Write(line);
            }

            sw.Close();
            fs.Close();
        }

        /// <summary>
        /// 查询数据
        /// </summary>
        /// <param name="sql"></param>
        /// <returns></returns>
        public DataSet GetDataSet(string sql)
        {
            DataSet dsResult = new DataSet();

            string sqlInsert = string.Format(sql);

            OracleConnection connection = new OracleConnection(ConnectionString);
            OracleCommand command = null;
            OracleTransaction transaction = null;

            try
            {
                connection.Open();
                transaction = connection.BeginTransaction();

                command = connection.CreateCommand();
                command.Transaction = transaction;

                command.CommandText = sqlInsert;


                OracleDataAdapter Adpt = new OracleDataAdapter(command);
                Adpt.Fill(dsResult, "NewTable");
                connection.Close();
            }
            catch (Exception ex)
            {
                connection.Close();
            }

            return dsResult;
        }

        /// <summary>
        /// 用药记录数据
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button5_Click(object sender, EventArgs e)
        {

            #region 用药记录数据

            string sql = @"

--附件6 用药记录
SELECT '01' P7501, --1 就诊类型  
       a.card_no P7502, --2 就诊卡号  
       a.reg_date P7506, --3 就诊日期  
       t.mo_order P7500, --4 顺序号  每条新增处方或医嘱的顺序号，用药记录唯一性标识 
       t.item_name P8016, --5 药物名称  
       1 P8017, --6药物使用频率
       t.qty P8018, --7药物使用总剂量 
       t.qty P8019, --8药物使用次剂量 
       t.dose_once P8020, --9药物使用剂量 
       t.oper_date P8021, --10药物使用开始时间  
       t.fee_date P8022 --11药物使用结束时间 
  FROM fin_opb_feedetail t, fin_opr_register a
 WHERE a.clinic_code = t.clinic_code
   and t.cancel_flag = '1'
   and t.pay_flag ='1'
   and t.qty > 0
   and t.fee_date>to_date('1900-01-01','yyyy-mm-dd')
   and t.clinic_code = '{0}'

UNION ALL

SELECT '03' P7501, --1 就诊类型  
       FII.Patient_No P7502, --2 就诊卡号  
       FII.in_DATE P7506, --3 就诊日期  
       o.mo_order P7500, --4 顺序号  每条新增处方或医嘱的顺序号，用药记录唯一性标识 
       o.item_name P8016, --5 药物名称  
       1 P8017, --6药物使用频率（日次数）
       o.qty_tot P8018, --7药物使用总剂量 
       o.dose_once P8019, --8药物使用次剂量  
       o.qty_tot P8020, --9药物使用剂量 
       case when o.date_bgn<to_date('1900-01-01','yyyy-mm-dd') then to_date('1800-01-01','yyyy-mm-dd') else o.date_bgn end P8021, --10药物使用开始时间 
       case when o.date_end<to_date('1900-01-01','yyyy-mm-dd') then to_date('1800-01-01','yyyy-mm-dd') else o.date_end end P8022 --11药物使用结束时间 
  FROM MET_IPM_ORDER o, fin_ipr_inmaininfo fii 
 WHERE 1 = 1 
   and o.confirm_flag = '1' 
   and o.qty_tot > 0
   AND fii.inpatient_no = o.inpatient_no
   and fii.inpatient_no ='{0}'




";

            #endregion

            #region 导出数据
            //获取流感患者
            DataSet vfp = FluDataSet;

            DataTable all = new DataTable();

            if (vfp.Tables[0].Rows.Count > 0)
            {

                DataTable dt = vfp.Tables[0];

                var query = from t in dt.AsEnumerable()
                            group t by new { t1 = t.Field<string>("clinic_code") } into m
                            select new
                            {
                                clinic_code = m.Key.t1
                            };
                if (query.ToList().Count > 0)
                {
                    query.ToList().ForEach(q =>
                    {
                        if (!q.clinic_code.Contains("-"))
                        {
                            DataSet ds = GetDataSet(string.Format(sql, q.clinic_code));

                            if (ds.Tables[0].Rows.Count > 0)
                            {
                                all.Merge(ds.Tables[0]);
                            }
                        }
                    });
                }

                if (all.Rows.Count > 0)
                {
                    string filePath = strFile + @"\pdr_" + System.DateTime.Now.ToString("yyyyMMddHHmmss") + ".csv";

                    TableToCsv(all, filePath);
                    //strZip = filePath.Replace("Excel", "ZipFile").Replace(".csv", ".zip");
                    strZip = filePath.Replace("Excel", "ZipFile");
                    

                    //压缩备份
                    string strZipBak = strZip.Replace("ZipFile", "ZipFileBak").Replace(".csv", ".zip");
                    ZipFile(strFile, strZipBak);

                    File.Copy(filePath, strZip, true);
                   
                }
            }
            #endregion

        }

        /// <summary>
        /// 出院流感病例数据
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button2_Click(object sender, EventArgs e)
        {

            #region 出院流感病例数据
            string sql = @"


SELECT 
  '45608783744060611A1001' AS P900,    --组织机构代码  指医疗机构执业许可证上面的机构代码。
  '佛山市顺德区北滘医院' AS P6891,    --医疗机构名称  指患者住院诊疗所在的医疗机构名称，按照《医疗机构执业许可证》登记的机构名称填写。
  '' P686,  
  '-' AS P800,    --健康卡号  在已统一发放[中华人民共和国居民健康卡]的地区填写健康卡号码，尚未发放[健康卡]的地区填写[-]
  M.xValue AS P1,    --医疗付费方式  值域范围参考RC032 
  CAST(A.FTIMES AS VARCHAR(800)) AS P2,    --住院次数  大于0的整数
  
  A.FPRN AS P3,    --病案号
  
  A.FNAME AS P4,    --姓名  
  A.FSEXBH AS P5,    --性别  值域范围参考RC001
  CONVERT(VARCHAR(20),A.FBIRTHDAY,20) AS P6,    --出生日期  格式 yyyy-MM-dd
  CAST(DATEDIFF(YEAR,A.FBIRTHDAY,FRYDATE) AS VARCHAR(3)) AS P7,    --年龄（岁）  患者入院年龄，指患者入院时按照日历计算的历法年龄，应以实足年龄的相应整数填写。大于或等于0的整数
  LEFT(A.FSTATUSBH,1) AS P8,    --婚姻  值域范围参考RC002
  A.FJOBBH AS P9,    --职业  值域范围参考RC003
   '' P101, --  生省份 
   '' P102, -- 出生地市
   '' P103,  --16 出生地县
  CASE WHEN LEFT(A.FNATIONALITYBH,1)='0' THEN RIGHT(A.FNATIONALITYBH,1) WHEN A.FNATIONALITYBH='99' THEN '66' ELSE A.FNATIONALITYBH END AS P11,    --民族  值域范围参考RC035
  A.FCOUNTRY AS P12,    --国籍  
  CASE isnull(dbo.fn_CheckIDCard(A.FIDCARD),'-') WHEN '' THEN '-' ELSE isnull(dbo.fn_CheckIDCard(A.FIDCARD),'-') end  AS P13,    --身份证号  住院患者入院时要如实填写15位或18位身份证号码
  

  CASE A.FCURRADDR WHEN '' THEN '-' ELSE isnull(A.FCURRADDR,'-') END AS P801,    --现住址  
    CASE isnull(A.FCURRTELE,'-')  WHEN '' THEN '-' ELSE isnull(A.FCURRTELE,'-') END AS P802,    --现住址电话  
    CASE A.FCURRPOST  WHEN '' THEN '-' ELSE isnull(A.FCURRPOST,'-') END AS P803,    --现住址邮政编码  6位数字
    CASE isnull(A.FDWADDR,'-')  WHEN '' THEN '-' ELSE isnull(A.FDWADDR,'-') END AS P14,    --工作单位及地址  
  CASE isnull(A.FDWTELE,'-')  WHEN '' THEN '-' ELSE isnull(A.FDWTELE,'-') end AS P15,    --工作单位电话  
  CASE A.FDWPOST  WHEN '' THEN '-' ELSE isnull(A.FDWPOST,'-') END AS P16,    --工作单位邮政编码  6位数字
  
  A.FHKADDR AS P17,    --户口地址  
    CASE A.FHKPOST WHEN '' THEN '-' ELSE isnull(A.FHKPOST,'-') END  AS P171,    --户口地址邮政编码  6位数字
  
  CASE isnull(A.FLXNAME,'-')  WHEN '' THEN '-' ELSE isnull(A.FLXNAME,'-') end AS P18,    --联系人姓名  
  K.xValue AS P19,    --联系人关系  值域范围参考RC033
  case A.FLXADDR WHEN '' THEN '-' ELSE isnull(A.FLXADDR,'-') END AS P20,    --联系人地址  
    A.FRYTJBH AS P804,    --入院途径  值域范围参考RC026
    CASE A.FLXTELE WHEN '' THEN '-' ELSE isnull(A.FLXTELE,'-') END AS P21,    --联系人电话 
  CASE WHEN A.FRYTIME BETWEEN '00' AND '23' THEN CONVERT(VARCHAR(20),A.FRYDATE,23) + ' ' + A.FRYTIME + ':00:00' ELSE CONVERT(VARCHAR(20),A.FRYDATE,20) END AS P22,    --入院时间  格式 yyyy-MM-dd HH:mm:ss；入院时间不能晚于出院时间
  H.xValue AS P23,    --入院科别  值域范围参考RC023
  '-' AS P231,    --入院病房  
  ISNULL(I.xValue,'-') AS P24,    --转科科别  值域范围参考RC023；转经多个科室时，值以英文逗号进行分隔
  CASE WHEN A.FCYTIME BETWEEN '00' AND '23' THEN CONVERT(VARCHAR(20),A.FCYDATE,23) + ' ' + A.FCYTIME + ':00:00' ELSE CONVERT(VARCHAR(20),A.FCYDATE,20) END AS P25,    --出院时间  格式 yyyy-MM-dd HH:mm:ss
  J.xValue AS P26,    --出院科别  值域范围参考RC023 
  '-' AS P261,    --出院病房  
  CAST(A.FDAYS AS VARCHAR(10)) AS P27,    --实际住院（天）  大于0整数。入院时间与出院时间只计算一天，例如：2018年6月12日入院，2018年6月15日出院，计住院天数为3天
  N.ICD AS P28,    --门（急）诊诊断编码  采用疾病分类代码国家临床版2.0编码（ICD-10）
  N.Diagnosis AS P281,    --门（急）诊诊断名称  采用疾病分类代码国家临床版2.0(ICD-10)与编码对应的诊断名称
  /*添加字段*/
    '' AS P29 , --入院时情况
  '' AS P30, -- 入院诊断编码  
  '' AS P301, --入院诊断名称
  '' AS P31,-- 入院后确诊日期
  /*添加字段*/
  P.ICD AS P321,    --出院主要诊断编码  采用疾病分类代码国家临床版2.0编码（ICD-10）
  P.Diagnosis AS P322,    --出院主要诊断名称  采用疾病分类代码国家临床版2.0(ICD-10)与编码对应的诊断名称
  B.FRYBQBH AS P805,    --出院主要诊断入院病情  值域范围参考RC027
  '' P323,
  '' P324, --其他诊断编码 1 
       '' P325, --其他诊断疾病描述 1
       '' P806, --其他诊断入院病情1
       '' P326, --其他诊断出院情况1 
       '' P327, --其他诊断编码 2
       '' P328, --其他诊断疾病描述 2
       '' P807, --其他诊断入院病情 2 
       '' P329, --其他诊断出院情况 2
       '' P3291, --其他诊断编码 3 
       '' P3292, --其他诊断疾病描述 3 
       '' P808, --其他诊断入院病情 3 
       '' P3293, --其他诊断出院情况 3 
       '' P3294, --其他诊断编码 4 
       '' P3295, --其他诊断疾病描述 4 
       '' P809, --其他诊断入院病情 4 
       '' P3296, --其他诊断出院情况 4 
       '' P3297, --其他诊断编码 5 
       '' P3298, --其他诊断疾病描述 5 
       '' P810, --其他诊断入院病情 5
       '' P3299, --其他诊断出院情况
       '' P3281, --其他诊断编码 6 
       '' P3282, --其他诊断疾病描述 6
       '' P811, --其他诊断入院病情 6
       '' P3283, --其他诊断出院情况 6
       '' P3284, --其他诊断编码 7 
       '' P3285, --其他诊断疾病描述 7 
       '' P812, --其他诊断入院病情 7 
       '' P3286, --其他诊断出院情况 7 
       '' P3287, --其他诊断编码 8 
       '' P3288, --其他诊断疾病描述 8 
       '' P813, --其他诊断入院病情 8 
       '' P3289, --其他诊断出院情况 8
       '' P3271, --其他诊断编码 9 
       '' P3272, --其他诊断疾病描述 9 
       '' P814, --其他诊断入院病情 9 
       '' P3273, --其他诊断出院情况 9 
       '' P3274, --其他诊断编码 10 
       '' P3275, --其他诊断疾病描述 10 
       '' P815, --其他诊断入院病情 10 
       '' P3276, --其他诊断出院情况 10
  '' P689,-- 医院感染总次数 
  '' P351, --病理诊断编码 1
       '' P352, --病理诊断名称 1
       '' P816, --病理号 1 
       '' P353, --病理诊断编码 2
       '' P354, --病理诊断名称 2
       '' P817, --病理号 2 
       '' P355, --病理诊断编码 3
       '' P356, --病理诊断名称 3
       '' P818, --病理号 3 
       
  isnull( D.FICDM,'') AS P361,    --损伤、中毒外部原因编码  采用疾病分类代码国家临床版2.0的编码(ICD-10)，主要诊断ICD编码首字母为S或T时必填
  isnull(D.FJBNAME,'') AS P362,    --损伤、中毒外部原因名称  采用疾病分类代码国家临床版2.0(ICD-10)编码对应的外部原因名称；主要诊断ICD编码首字母为S或T时必填
      '' P363, --损伤、中毒的外 部因素编码 2 
       '' P364, --损伤、中毒的外部因素名称 2 
       '' P365, --损伤、中毒的外 部因素编码 3
       '' P366, --损伤、中毒的外部因素名称 3 
       '' P371, --过敏源
       '' P372, --过敏药物名称
       '' P38, --HBsAg
       '' P39, --HCV-Ab
       '' P40, --HIV-Ab 
       '' P411, --门诊与出院诊断 符合情况 
       '' P412, --入院与出院诊断 符合情况 
       '' P413, --术前与术后诊断 符合情况 
       '' P414, --临床与病理诊断 符合情况 
       '' P415, --放射与病理诊断 符合情况
       '' P421, --抢救次数
       '' P422, --抢救成功次数
       '' P687, --最高诊断依据
       '' P688, --分化程度
  A.FKZR AS P431,    --科主任  
  A.FZRDOCTOR AS P432,    --主（副主）任医师 
  A.FZZDOCT AS P433,    --主治医师  
  A.FZYDOCT AS P434,    --住院医师  
    CASE A.FNURSE WHEN '' THEN '-' ELSE isnull(A.FNURSE,'-') END  AS P819,    --责任护士
  CASE isnull(A.FJXDOCT,'无') WHEN '无' THEN '' ELSE A.FJXDOCT end AS P435,    --进修医师
  CASE isnull(A.FSXDOCT,'无') WHEN '无' THEN '' ELSE A.FSXDOCT end AS P436,    --实习医师  
  '' P437,
  A.FBMY AS P438,    --编码员  
  
  A.FQUALITYBH AS P44,    --病案质量  值域范围参考RC011
  A.FZKDOCT AS P45,    --质控医师  
  A.FZKNURSE AS P46,    --质控护师  
  CONVERT(VARCHAR(10),A.FZKRQ,23) AS P47,    --质控日期  格式 yyyy-MM-dd
  
  isnull( O.OpCode,'-') AS P490,    --主要手术操作编码  手术操作名称第一行为[主要手术操作]，采用手术操作分类代码国家临床版2.0编码（ICD-9-CM3）
  isnull(CONVERT(VARCHAR(20),C.FOPDATE,20),'-') AS P491,    --主要手术操作日期  格式 yyyy-MM-dd HH:mm:ss
  isnull( C.FSSJBBH,'') AS P820,    --主要手术操作级别  手术及操作编码属性为手术或介入治疗代码时必填。值域范围参考RC029。
  isnull( O.OpName,'-') AS P492,    --主要手术操作名称  手术操作名称第一行为[主要手术操作]，采用手术操作分类代码国家临床版2.0（ICD-9-CM3）编码对应的名称 
  '' P493,
  '' AS P494, --主要手术持续时间
  isnull( C.FDOCNAME,'') AS P495,    --主要手术操作术者  手术及操作编码属性为手术或介入治疗代码时必填
  isnull( C.FOPDOCT1,'') AS P496,    --主要手术操作Ⅰ助  手术及操作编码属性为手术或介入治疗代码时必填
  isnull( C.FOPDOCT2,'') AS P497,    --主要手术操作Ⅱ助  手术及操作编码属性为手术或介入治疗代码时必填
  isnull( G.xValue,'') AS P498,    --主要手术操作麻醉方式  手术编码属性为手术时必填，值域范围参考RC013 
  '' AS P4981, --主要手术麻醉分级
  isnull( L.xValue,'') AS P499,    --主要手术操作切口愈合等级  手术编码属性为手术时必填，值域范围参考RC014
  isnull( C.FMZDOCT,'') AS P4910,    --主要手术操作麻醉医师  手术及操作编码属性为手术时必填

       '' P4911, --手术操作编码 2
       '' P4912, --手术操作日期 2
       '' P821, --手术级别 2
       '' P4913, --手术操作名称 2 
       '' P4914, --手术操作部位 2 
       '' P4915, --手术持续时间 2
       '' P4916, --术者 2
       '' P4917, --Ⅰ助 2
       '' P4918, --Ⅱ助 2
       '' P4919, --麻醉方式 2
       '' P4982, --麻醉分级 2
       '' P4920, --切口愈合等级 2
       '' P4921, --麻醉医师 2 
       '' P4922, --手术操作编码 3
       '' P4923, --手术操作日期 3
       '' P822, --手术级别 3
       '' P4924, --手术操作名称 3 
       '' P4925, --手术操作部位 3 
       '' P4526, --手术持续时间 3
       '' P4527, --术者 3
       '' P4528, --Ⅰ助 3
       '' P4529, --Ⅱ助 3
       '' P4530, --麻醉方式 3
       '' P4983, --麻醉分级 3
       '' P4531, --切口愈合等级 3
       '' P4532, --麻醉医师 3 
       '' P4533, --手术操作编码 4
       '' P4534, --手术操作日期 4
       '' P823, --手术级别 4
       '' P4535, --手术操作名称 4 
       '' P4536, --手术操作部位 4 
       '' P4537, --手术持续时间 4
       '' P4538, --术者 4
       '' P4539, --Ⅰ助 4
       '' P4540, --Ⅱ助 4
       '' P4541, --麻醉方式 4
       '' P4984, --麻醉分级 4
       '' P4542, --切口愈合等级 4
       '' P4543, --麻醉医师 4 
       '' P4544, --手术操作编码 5
       '' P4545, --手术操作日期 5
       '' P824, --手术级别 5
       '' P4546, --手术操作名称 5 
       '' P4547, --手术操作部位 5 
       '' P4548, --手术持续时间 5
       '' P4549, --术者 5
       '' P4550, --Ⅰ助 5
       '' P4551, --Ⅱ助 5
       '' P4552, --麻醉方式 5
       '' P4985, --麻醉分级 5
       '' P4553, --切口愈合等级 5
       '' P4554, --麻醉医师 5 
       '' P45002, --手术操作编码 6
       '' P45003, --手术操作日期 6
       '' P825, --手术级别 6
       '' p45004, --手术操作名称 6 
       '' p45005, --手术操作部位 6 
       '' p45006, --手术持续时间 6
       '' p45007, --术者 6
       '' p45008, --Ⅰ助 6
       '' p45009, --Ⅱ助 6
       '' p45010, --麻醉方式 6
       '' p45011, --麻醉分级 6
       '' p45012, --切口愈合等级 6
       '' p45013, --麻醉医师 6 
       '' p45014, --手术操作编码 7
       '' p45015, --手术操作日期 7
       '' P826, --手术级别 7
       '' p45016, --手术操作名称 7 
       '' p45017, --手术操作部位 7 
       '' p45018, --手术持续时间 7
       '' p45019, --术者 7
       '' p45020, --Ⅰ助 7
       '' p45021, --Ⅱ助 7
       '' p45022, --麻醉方式 7
       '' p45023, --麻醉分级 7
       '' p45024, --切口愈合等级 7
       '' p45025, --麻醉医师 7 
       '' p45026, --手术操作编码 8
       '' p45027, --手术操作日期 8
       '' P827, --手术级别 8
       '' p45028, --手术操作名称 8 
       '' p45029, --手术操作部位 8 
       '' p45030, --手术持续时间 8
       '' p45031, --术者 8
       '' p45032, --Ⅰ助 8
       '' p45033, --Ⅱ助 8
       '' p45034, --麻醉方式 8  
       '' p45035, --麻醉分级 8
       '' p45036, --切口愈合等级 8
       '' p45037, --麻醉医师 8 
       '' p45038, --手术操作编码 9
       '' p45039, --手术操作日期 9
       '' P828, --手术级别 9
       '' p45040, --手术操作名称 9 
       '' p45041, --手术操作部位 9 
       '' p45042, --手术持续时间 9
       '' p45043, --术者 9
       '' p45044, --Ⅰ助 9
       '' p45045, --Ⅱ助 9
       '' p45046, --麻醉方式 9
       '' p45047, --麻醉分级 9
       '' p45048, --切口愈合等级 9
       '' p45049, --麻醉医师 9 
       '' p45050, --手术操作编码 10
       '' p45051, --手术操作日期 10
       '' P829, --手术级别 10
       '' p45052, --手术操作名称 10 
       '' p45053, --手术操作部位 10 
       '' p45054, --手术持续时间 10
       '' p45055, --术者 10
       '' p45056, --Ⅰ助 10
       '' p45057, --Ⅱ助 10
       '' p45058, --麻醉方式 10
       '' p45059, --麻醉分级 10
       '' p45060, --切口愈合等级 10
       '' p45061, --麻醉医师 10
       '' P561, --特级护理天数
       '' P562, --一级护理天数
       '' P563, --二级护理天数
       '' P564, --三级护理天数 
       '' P6911, --重症监护室名称1
       '' P6912, --进入时间 1
       '' P6913, --退出时间 1 
       '' P6914, --重症监护室名称2
       '' P6915, --进入时间 2
       '' P6916, --退出时间 2 
       '' P6917, --重症监护室名称3
       '' P6918, --进入时间 3
       '' P6919, --退出时间 3 
       '' P6920, --重症监护室名称4
       '' P6921, --进入时间 4
       '' P6922, --退出时间 4 
       '' P6923, --重症监护室名称5
       '' P6924, --进入时间 5
       '' P6925, --退出时间 5 
    A.FBODYBH AS P57,    --死亡患者尸检  值域范围参考RC016
   '' P58, --手术、治疗、检 查、诊断为本院第一例 
       '' P581, --手术患者类型
       '' P60, --随诊
       '' P611, --随诊周数
       '' P612, --随诊月数
       '' P613, --随诊年数
       '' P59, --示教病例
   A.FBLOODBH AS P62,    --ABO血型  值域范围参考RC030
    CASE A.FRHBH WHEN '' THEN '-' ELSE isnull(A.FRHBH,'-') END  P63,
   '' P64, --输血反应 
       '' P651, --红细胞 
       '' P652, --血小板
       '' P653, --血浆
       '' P654, --全血
       '' P655, --自体回收
       '' P656, --其它  
   CASE WHEN DATEDIFF(YEAR,FBIRTHDAY,FRYDATE)=0 THEN DATEDIFF(DAY,FBIRTHDAY,FRYDATE)+1 ELSE 0 END AS P66,    --年龄不足1周岁的年龄（天）  按照实足年龄的天数填写。年龄不足1周岁时填写，年龄值A14应为0，取值范围：大于或等于0小于365，入院时间减出生日期后取整数，不足一天按0天计算。
  CASE A.FCSTZ WHEN 0 THEN '-' ELSE CAST(A.FCSTZ AS VARCHAR) END AS P681,    --新生儿出生体重(克)  测量新生儿体重要求精确到10克；应在活产后一小时内称取重量。1、产妇和新生儿病案填写，从出生到28天为新生儿期，双胎及以上不同胎儿体重则继续填写下面的新生儿出生体重。2、新生儿体重范围：100克-9999克，产妇的主要诊断或其他诊断编码中含有Z37.0,Z37.2, Z37.3, Z37.5, Z37.6编码时，必须填写新生儿出生体重
  '' AS P682,    --新生儿出生体重(克)2  新生儿体重范围：100克-9999克
  '' AS P683,    --新生儿出生体重(克)3  新生儿体重范围：100克-9999克
  '' AS P684,    --新生儿出生体重(克)4  新生儿体重范围：100克-9999克
  '' AS P685,    --新生儿出生体重(克)5  新生儿体重范围：100克-9999克
    CASE A.FRYTZ WHEN 0 THEN '-' ELSE CAST(A.FRYTZ AS VARCHAR) END AS P67,    --新生儿入院体重（克）  指新生儿入院当日体重，100克-9999克，精确到10克；[新生儿入院体重]与[年龄不足1周岁的年龄（天）]互为逻辑校验项，小于等于28天的新生儿必填。
  CASE A.FRYQHMHOURS WHEN '' THEN 0 ELSE isnull(A.FRYQHMHOURS,0) END AS P731,    --颅脑损伤患者入院前昏迷时间(小时)  大于等于0，小于24整数。
  CASE A.FRYQHMMINS WHEN '' THEN 0 ELSE isnull(A.FRYQHMMINS,0)  END AS P732,    --颅脑损伤患者入院前昏迷时间(分钟)  大于等于0，小于60整数。
  CASE A.FRYHMHOURS WHEN '' THEN 0 ELSE isnull(A.FRYHMHOURS,0) END AS P733,    --颅脑损伤患者入院后昏迷时间(小时)  大于等于0，小于24整数。
  CASE A.FRYHMMINS WHEN '' THEN 0 ELSE isnull(A.FRYHMMINS,0) END AS P734,    --颅脑损伤患者入院后昏迷时间(分钟)  大于等于0，小于60整数。
  '' P72,
  CASE A.FISAGAINRYBH  WHEN '' THEN '-' ELSE isnull(a.FISAGAINRYBH,'-') END AS P830,    --是否有出院31日内再住院计划  值域范围参考RC028；指患者本次住院出院后31天内是否有诊疗需要的再住院安排。如果有再住院计划，则需要填写目的，如：进行二次手术
  A.FISAGAINRYMD AS P831,    --出院31天再住院计划目的  是否有出院31日内再住院计划填[有]时必填
  A.FLYFSBH AS P741,    --离院方式  值域范围参考RC019；指患者本次住院出院的方式，填写相应的阿拉伯数字
  CASE A.FLYFSBH WHEN '2' THEN A.FYZOUTHOSTITAL ELSE NULL END AS P742,    --医嘱转院、转社区卫生服务机构/乡镇卫生院名称  离院方式为医嘱转院或医嘱转社区患者必填
  CASE A.FLYFSBH WHEN '3' THEN A.FSQOUTHOSTITAL ELSE NULL END AS P743,    --医嘱转院、转社区卫生服务机构/乡镇卫生院名称  离院方式为医嘱转院或医嘱转社区患者必填
  CASE WHEN A.FQTF<0 THEN A.FSUM1-A.FQTF ELSE A.FSUM1 END AS P782,    --住院总费用  住院总费用必填且大于0；总费用大于或等于分项费用之和；
  A.FZFJE AS P751,    --住院总费用其中自付金额  小于等于总费用
  A.FZHFWLYLF AS P752,    --1.一般医疗服务费  
  A.FZHFWLCZF AS P754,    --2.一般治疗操作费  
  A.FZHFWLHLF AS P755,    --3.护理费  
  A.FZHFWLQTF AS P756,    --4.综合医疗服务类其他费用  
  A.FZDLBLF AS P757,    --5.病理诊断费  
  A.FZDLSSSF AS P758,    --6.实验室诊断费  
  A.FZDLYXF AS P759,    --7.影像学诊断费  
  A.FZDLLCF AS P760,    --8.临床诊断项目费  
  A.FZLLFFSSF AS P761,    --9.非手术治疗项目费  
    A.FZLLFWLZWLF AS P762,    --其中：临床物理治疗费  
  A.FZLLFSSF AS P763,    --10.手术治疗费  
  A.FZLLFMZF AS P764,    --其中：麻醉费  
  A.FZLLFSSZLF AS P765,    --其中：手术费  
  A.FKFLKFF AS P767,    --11.康复费  
  A.FZYLZF AS P768,    --12.中医治疗费  
  A.FXYF AS P769,    --13.西药费  
  A.FXYLGJF AS P770,    --其中：抗菌药物费  
  A.FZCHYF AS P771,    --14.中成药费  
  A.FZCYF AS P7702,    --15.中草药费  
  A.FXYLXF AS P773,    --16.血费  
  A.FXYLBQBF AS P774,    --17.白蛋白类制品费  
  A.FXYLQDBF AS P775,    --18.球蛋白类制品费  
  A.FXYLYXYZF AS P776,    --19.凝血因子类制品费  
  A.FXYLXBYZF AS P777,    --20.细胞因子类制品费  
  A.FHCLCJF AS P778,    --21.检查用一次性医用材料费  
  A.FHCLZLF AS P779,    --22.治疗用一次性医用材料费  
  A.FHCLSSF AS P780,    --23.手术用一次性医用材料费  
  CASE WHEN A.FQTF<0 THEN 0 ELSE A.FQTF END AS P781    --24.其他费：
  FROM TPATIENTVISIT A WITH(NOLOCK) LEFT JOIN TDIAGNOSE B WITH(NOLOCK) ON A.FPRN=B.FPRN AND A.FTIMES=B.FTIMES AND FZDLX='1'
    LEFT JOIN TOPERATION C WITH(NOLOCK) ON A.FPRN=C.FPRN AND A.FTIMES=C.FTIMES AND C.FPX=1
    LEFT JOIN TDIAGNOSE D WITH(NOLOCK) ON A.FPRN=D.FPRN AND A.FTIMES=D.FTIMES AND D.FZDLX='s' 
    LEFT JOIN xRef_ICD F ON A.FPHZDBH=F.FICDM AND F.Category='M码'
    LEFT JOIN xRef_Dict E ON LEFT(A.FHKADDR,2)=E.ItemCode AND E.Dict='RC036'
    LEFT JOIN xRef_Dict G ON C.FMAZUIBH=G.ItemCode AND G.Dict='RC013'
    LEFT JOIN xRef_Dept H ON A.FRYTYKH=H.ItemCode AND H.Dict='RC023'
    LEFT JOIN xRef_Dept I ON A.FZKTYKH=I.ItemCode AND I.Dict='RC023'
    LEFT JOIN xRef_Dept J ON A.FCYTYKH=J.ItemCode AND J.Dict='RC023'
    LEFT JOIN xRef_Dict K ON A.FRELATE=K.ItemCode AND K.Dict='RC033'    
    LEFT JOIN xRef_Dict L ON ISNULL(C.FQIEKOUBH,'1')+ISNULL(C.FYUHEBH,'0')=L.ItemCode AND L.Dict='RC014'
    LEFT JOIN xRef_Dict M ON CASE WHEN A.FFBBHNEW BETWEEN '01' AND '03' THEN A.FSOURCEBH+A.FFBBHNEW ELSE A.FFBBHNEW END=M.ItemCode AND M.Dict='RC032'
    LEFT JOIN xRef_ICD N ON A.FMZZDBH=N.FICDM AND N.Category='ICD10'
    LEFT JOIN xRef_ICD9CM3 O ON C.FOPCODE=O.FOPCODE
    LEFT JOIN xRef_ICD P ON B.FICDM=P.FICDM AND P.Category='ICD10'
  WHERE A.FPRN ='{0}' AND A.FCYDATE ='{1}'

            ";
            #endregion

            #region 导出数据

            //获取流感患者
            DataSet vfp = FluDataSet;//

            DataTable all = new DataTable();

            if (vfp.Tables[0].Rows.Count > 0)
            {

                DataTable dt = vfp.Tables[0];

                DataView receiveDV = dt.DefaultView;

                receiveDV.RowFilter = "patient_type ='住院'";
                dt = receiveDV.ToTable();

                var query = from t in dt.AsEnumerable()
                            group t by new { t1 = t.Field<string>("clinic_code") } into m
                            select new
                            {
                                clinic_code = m.Key.t1
                            };
                if (query.ToList().Count > 0)
                {
                    query.ToList().ForEach(q =>
                    {
                        if (!q.clinic_code.Contains("-"))
                        {

                            DataSet ds = GetDataSet(string.Format(sql, q.clinic_code));

                            if (ds.Tables[0].Rows.Count > 0)
                            {
                                all.Merge(ds.Tables[0]);
                            }
                        }
                    });
                }

                if (all.Rows.Count > 0)
                {
                    string filePath = strFile + @"\hqms_" + System.DateTime.Now.ToString("yyyyMMddHHmmss") + ".csv";

                    TableToCsv(all, filePath);
                    strZip = filePath.Replace("Excel", "ZipFile");


                    //压缩备份
                    string strZipBak = strZip.Replace("ZipFile", "ZipFileBak").Replace(".csv", ".zip");
                    ZipFile(strFile, strZipBak);

                    File.Copy(filePath, strZip, true);
                }
            }
            #endregion

        }

        /// <summary>
        ///  删除文件
        /// </summary>
        /// <param name="theDir"></param>
        /// <param name="nLevel"></param>
        /// <param name="Rn"></param>
        /// <returns></returns>
        private static void RemoveFiles(string dirFile)//递归目录 文件 
        {
            DirectoryInfo theFolder = new DirectoryInfo(dirFile);

            FileInfo[] fileInfo = theFolder.GetFiles();
            foreach (FileInfo NextFile in fileInfo)  //遍历文件
            {
                System.IO.File.Delete(NextFile.DirectoryName + "\\" + NextFile.Name); //移动旧文件                      

            }
        }

        /// <summary>
        ///  压缩文件夹
        /// </summary>
        /// <param name="strFile">待压缩文件夹的路径</param>
        /// <param name="strZip">压缩zip的路径</param>
        private bool ZipFile(string strFile, string strZip)
        {
            bool flag = false;
            if (strFile[strFile.Length - 1] != Path.DirectorySeparatorChar)
            {
                strFile += Path.DirectorySeparatorChar;
            }
            ZipOutputStream s = null;
            try
            {
                s = new ZipOutputStream(System.IO.File.Create(strZip));
                s.SetLevel(6); // 压缩等级
                flag = zip(strFile, s, strFile);//递归压缩文件夹
            }
            catch (Exception e)
            {
                throw e;
            }
            finally
            {
                s.Finish();
                s.Close();
            }
            return flag;
        }

        /// <summary>
        /// 压缩文件夹的递归方法
        /// </summary>
        /// <param name="strFile">待压缩文件夹的路径</param>
        /// <param name="s"></param>
        /// <param name="staticFile"></param>
        /// <returns></returns>
        private bool zip(string strFile, ZipOutputStream s, string staticFile)
        {
            bool flag = false;
            if (strFile[strFile.Length - 1] != Path.DirectorySeparatorChar)
            {
                strFile += Path.DirectorySeparatorChar;
            }
            //文件检验码
            Crc32 crc = new Crc32();
            //当前文件夹的所有文件的绝对路径
            string[] filenames = Directory.GetFileSystemEntries(strFile);
            foreach (string file in filenames)
            {
                //如果待压缩的文件夹内存在子文件夹，递归调用
                if (Directory.Exists(file))
                {
                    zip(file, s, staticFile);
                }
                else // 否则直接压缩文件
                {
                    //打开压缩文件
                    FileStream fs = System.IO.File.OpenRead(file);
                    //设置内存缓冲区大小
                    byte[] buffer = new byte[fs.Length];
                    //创建一个内存缓冲区
                    fs.Read(buffer, 0, buffer.Length);
                    //以待文件夹的名称做为压缩文件名
                    string tempfile = file.Substring(staticFile.LastIndexOf("\\") + 1);
                    //创建压缩文件
                    ZipEntry entry = new ZipEntry(tempfile);
                    entry.DateTime = DateTime.Now;
                    entry.Size = fs.Length;
                    fs.Close();
                    crc.Reset();
                    crc.Update(buffer);
                    entry.Crc = crc.Value;
                    s.PutNextEntry(entry);
                    s.Write(buffer, 0, buffer.Length);
                }
            }
            return flag;
        }

        /// <summary>
        /// 出院小结数据
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button3_Click(object sender, EventArgs e)
        {

            #region 出院小结数据
            string sql = @"
SELECT INFO.INPATIENT_NO AS P3, --病案号
       INFO.NAME AS P4, --姓名
       DECODE(INFO.SEX_CODE, 'F', '2', 'M', '1', '9') AS P5, --性别
       (CASE
         WHEN ADD_MONTHS(INFO.BIRTHDAY,
                         (EXTRACT(YEAR FROM INFO.IN_DATE) -
                         EXTRACT(YEAR FROM INFO.BIRTHDAY)) * 12) <
              INFO.IN_DATE THEN
          EXTRACT(YEAR FROM INFO.IN_DATE) - EXTRACT(YEAR FROM INFO.BIRTHDAY)
         ELSE
          EXTRACT(YEAR FROM INFO.IN_DATE) - EXTRACT(YEAR FROM INFO.BIRTHDAY) - 1
       END) AS P7, --年龄
       INFO.IN_DATE AS P22, --入院日期
       (SELECT (SELECT A.INTERFACE_CODE
                  FROM COM_COMPARE_INTERFACEINFO A
                 WHERE INFO.DEPT_CODE = A.LOCAL_CODE
                   AND A.INTERFACE_TYPE = '科室对照'
                   AND ROWNUM = 1)
          FROM COM_SHIFTDATA CS
         WHERE CS.SHIFT_TYPE = 'B'
           AND CS.CLINIC_NO = INFO.INPATIENT_NO
           AND CS.HAPPEN_NO = (SELECT MAX(CSA.HAPPEN_NO)
                                 FROM COM_SHIFTDATA CSA
                                WHERE CSA.CLINIC_NO = CS.CLINIC_NO
                                  AND CSA.SHIFT_TYPE = 'B')
           AND ROWNUM = 1) AS P23, --入院科别
       (SELECT (SELECT A.INTERFACE_CODE
                  FROM COM_COMPARE_INTERFACEINFO A
                 WHERE INFO.DEPT_CODE = A.LOCAL_CODE
                   AND A.INTERFACE_TYPE = '科室对照'
                   AND ROWNUM = 1)
          FROM COM_SHIFTDATA CS
         WHERE CS.SHIFT_TYPE = 'RO'
           AND CS.CLINIC_NO = INFO.INPATIENT_NO
           AND CS.HAPPEN_NO = (SELECT MAX(CSA.HAPPEN_NO)
                                 FROM COM_SHIFTDATA CSA
                                WHERE CSA.CLINIC_NO = CS.CLINIC_NO
                                  AND CSA.SHIFT_TYPE = 'RO')
           AND ROWNUM = 1) AS P24, --转科科别
       INFO.OUT_DATE AS P25, --出院日期
       (SELECT A.INTERFACE_CODE
          FROM COM_COMPARE_INTERFACEINFO A
         WHERE INFO.DEPT_CODE = A.LOCAL_CODE
           AND A.INTERFACE_TYPE = '科室对照'
           AND ROWNUM = 1) AS P26, ---出院科别
       TRUNC(INFO.OUT_DATE - INFO.IN_DATE) AS P27, --实际住院天数
       NVL(INFO.CLINIC_DIAGNOSE,
           (SELECT M.DIAG_NAME
              FROM MET_CAS_DIAGNOSE M
             WHERE M.INPATIENT_NO = INFO.INPATIENT_NO
               AND m.happen_no = 1)) AS P8600, --入院诊断
       NVL((SELECT M.DIAG_NAME
             FROM MET_CAS_DIAGNOSE M
            WHERE M.INPATIENT_NO = INFO.INPATIENT_NO
              AND m.happen_no = 1),
           '-') AS P8601, ----出院诊断
       '' AS P8602, -------入院情况及诊疗经过
       '' AS P8603, --------出院情况及治疗结果
       '' AS P8604 -----出院医嘱
  FROM FIN_IPR_INMAININFO INFO
 WHERE INFO.IN_STATE IN ('B', 'O')
   AND INFO.INPATIENT_NO = '{0}'";

            #endregion

            #region 导出数据
            //获取流感患者
            DataSet vfp = FluDataSet;

            DataTable all = new DataTable();

            if (vfp.Tables[0].Rows.Count > 0)
            {
                DataTable dt = vfp.Tables[0];

                DataView receiveDV = dt.DefaultView;

                receiveDV.RowFilter = "patient_type ='住院'";
                dt = receiveDV.ToTable();

                var query = from t in dt.AsEnumerable()
                            group t by new { t1 = t.Field<string>("clinic_code") } into m
                            select new
                            {
                                clinic_code = m.Key.t1
                            };
                if (query.ToList().Count > 0)
                {
                    query.ToList().ForEach(q =>
                    {
                        if (!q.clinic_code.Contains("-"))
                        {
                            DataSet ds = GetDataSet(string.Format(sql, q.clinic_code));

                            if (ds.Tables[0].Rows.Count > 0)
                            {
                                all.Merge(ds.Tables[0]);
                            }
                        }
                    });
                }

                if (all.Rows.Count > 0)
                {
                    string filePath = strFile + @"\hda_" + System.DateTime.Now.ToString("yyyyMMddHHmmss") + ".csv";

                    TableToCsv(all, filePath);
                    //strZip = filePath.Replace("Excel", "ZipFile").Replace(".csv", ".zip");
                    strZip = filePath.Replace("Excel", "ZipFile");
                    

                    //压缩备份
                    string strZipBak = strZip.Replace("ZipFile", "ZipFileBak").Replace(".csv", ".zip");
                    ZipFile(strFile, strZipBak);

                    File.Copy(filePath, strZip, true);
                }
            }
            #endregion

        }

        /// <summary>
        /// 获取流感患者信息
        /// </summary>
        /// <returns></returns>
        public DataSet GetFluUserData()
        {
            string sTime = System.DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd");
            string eTime = System.DateTime.Now.ToString("yyyy-MM-dd");

            if (this.txtSartTime.Text.Trim() != "")
            {
                sTime = this.txtSartTime.Text.Trim();
            }

            if (this.txtEndTime.Text.Trim() != "")
            {
                eTime = this.txtEndTime.Text.Trim();
            }

            string sqlVfp = string.Format(@"

SELECT DISTINCT CLINIC_CODE, PATIENT_TYPE
  FROM (SELECT C.CLINIC_CODE CLINIC_CODE, '门诊' PATIENT_TYPE
          FROM FIN_OPR_REGISTER C, MET_CAS_DIAGNOSE D
         WHERE C.CLINIC_CODE = D.INPATIENT_NO
           AND SUBSTR(D.ICD_CODE, 0, 3) BETWEEN 'J00' AND 'J99'
           AND C.VALID_FLAG = '1'
           AND D.IS30DISEASE = '1'
           AND C.REG_DATE > TO_DATE('{0} 00:00:00', 'yyyy-mm-dd hh24:mi:ss')
           AND C.REG_DATE <=
               TO_DATE('{1} 23:59:59', 'yyyy-mm-dd hh24:mi:ss')
        --根据门诊诊断名称
        /*UNION ALL
        SELECT DISTINCT B.INPATIENT_NO CLINIC_CODE, '门诊' PATIENT_TYPE
          FROM MET_CAS_DIAGNOSE B,fin_opr_register r
         WHERE (B.DIAG_NAME LIKE '甲%流' OR B.DIAG_NAME LIKE '%乙%流%' OR
               B.DIAG_NAME LIKE '%流%感%')
           AND B.VALID_FLAG = '1'
           and b.clinic_code = r.clinic_code
           AND r.reg_date > TO_DATE('{0} 00:00:00', 'yyyy-mm-dd hh24:mi:ss')
           AND r.reg_date <= TO_DATE('{1} 23:59:59', 'yyyy-mm-dd hh24:mi:ss')*/
        UNION ALL
        --根据门诊药品名称       
        SELECT B.CLINIC_CODE, '门诊' PATIENT_TYPE
          FROM FIN_OPR_REGISTER A, FIN_OPB_FEEDETAIL B
         WHERE A.CLINIC_CODE = B.CLINIC_CODE
           AND A.REG_DATE > TO_DATE('{0} 00:00:00', 'yyyy-mm-dd hh24:mi:ss')
           AND A.REG_DATE <=
               TO_DATE('{1} 23:59:59', 'yyyy-mm-dd hh24:mi:ss')
           AND A.VALID_FLAG = '1'
           AND (INSTR(B.ITEM_NAME, '奥司他韦') > 0 OR
               INSTR(B.ITEM_NAME, '扎那米韦') > 0 OR
               INSTR(B.ITEM_NAME, '阿比多尔') > 0 OR
               INSTR(B.ITEM_NAME, '阿比朵尔') > 0 OR
               INSTR(B.ITEM_NAME, '莲花清瘟') > 0 OR
               INSTR(B.ITEM_NAME, '金花清感颗粒') > 0 OR
               INSTR(B.ITEM_NAME, '金刚烷胺') > 0 OR
               INSTR(B.ITEM_NAME, '金刚乙胺') > 0 OR
               INSTR(B.ITEM_NAME, '利巴韦林') > 0)
           AND B.PAY_FLAG = '1'
           AND B.CANCEL_FLAG = '1'
        UNION ALL
        --门诊病历主诉提取
        SELECT V.CLINIC_CODE, '门诊' PATIENT_TYPE
          FROM FIN_OPR_REGISTER V, MET_CAS_HISTORY CURE
         WHERE V.CLINIC_CODE = CURE.CLINIC_CODE
           AND V.VALID_FLAG = '1'
           AND CURE.CASEMAIN IS NOT NULL
           AND (CURE.CASEMAIN LIKE '%甲%流' OR CURE.CASEMAIN LIKE '%乙%流%' OR
               CURE.CASEMAIN LIKE '%流%感%' OR CURE.CASEMAIN LIKE '%咳嗽%' OR
               CURE.CASEMAIN LIKE '%发烧%' OR CURE.CASEMAIN LIKE '%高热%')
           AND CURE.OPER_DATE >
               TO_DATE('{0} 00:00:00', 'yyyy-mm-dd hh24:mi:ss')
           AND CURE.OPER_DATE <=
               TO_DATE('{1} 23:59:59', 'yyyy-mm-dd hh24:mi:ss')
        UNION ALL
        -- 住院诊断
        SELECT B.INPATIENT_NO CLINIC_CODE, '住院' PATIENT_TYPE
          FROM MET_CAS_DIAGNOSE A, FIN_IPR_INMAININFO B
         WHERE A.INPATIENT_NO = B.INPATIENT_NO
           AND ((INSTR(A.DIAG_NAME, '甲') > 0 AND
               INSTR(A.DIAG_NAME, '流') > INSTR(A.DIAG_NAME, '甲') AND
               A.DIAG_NAME NOT LIKE '%流感嗜血%' AND
               A.DIAG_NAME NOT LIKE '%副流感%' AND
               A.DIAG_NAME NOT LIKE '%血流感染%') OR
               (INSTR(A.DIAG_NAME, '乙') > 0 AND
               INSTR(A.DIAG_NAME, '流') > INSTR(A.DIAG_NAME, '乙') AND
               A.DIAG_NAME NOT LIKE '%流感嗜血%' AND
               A.DIAG_NAME NOT LIKE '%副流感%' AND
               A.DIAG_NAME NOT LIKE '%血流感染%') OR
               (INSTR(A.DIAG_NAME, '流') > 0 AND
               INSTR(A.DIAG_NAME, '感') > INSTR(A.DIAG_NAME, '流') AND
               A.DIAG_NAME NOT LIKE '%流感嗜血%' AND
               A.DIAG_NAME NOT LIKE '%副流感%' AND
               A.DIAG_NAME NOT LIKE '%血流感染%') OR
               (INSTR(A.DIAG_NAME, 'H') > 0 AND
               INSTR(A.DIAG_NAME, 'N') > INSTR(A.DIAG_NAME, 'H')) OR
               (INSTR(A.DIAG_NAME, '高热') > 0 OR
               INSTR(A.DIAG_NAME, '发烧') > 0 OR
               INSTR(A.DIAG_NAME, '发热') > 0) AND
               (INSTR(A.DIAG_NAME, '咳嗽') > 0 OR
               INSTR(A.DIAG_NAME, '咳痰') > 0) OR
               (A.ICD_CODE IN (SELECT A.ICD_CODE
                                  FROM MET_COM_ICD10 A
                                 WHERE A.ICD_CODE > 'J00'
                                   AND A.ICD_CODE < 'J99'
                                   AND A.VALID_STATE = '1')) -- 通过诊断编码
               )
           AND B.IN_DATE > TO_DATE('{0} 00:00:00', 'yyyy-mm-dd hh24:mi:ss')
           AND B.IN_DATE <= TO_DATE('{1} 23:59:59', 'yyyy-mm-dd hh24:mi:ss')
           AND B.IN_STATE = 'I'
        UNION ALL
        --住院使用了这个药品的患者
        SELECT AF.INPATIENT_NO CLINIC_CODE, '住院' PATIENT_TYPE
          FROM FIN_IPR_INMAININFO AF, FIN_IPB_MEDICINELIST BD
         WHERE AF.INPATIENT_NO = BD.INPATIENT_NO
           AND AF.IN_DATE > TO_DATE('{0} 00:00:00', 'yyyy-mm-dd hh24:mi:ss')
           AND AF.IN_DATE <=
               TO_DATE('{1} 23:59:59', 'yyyy-mm-dd hh24:mi:ss')
           AND BD.TRANS_TYPE = '1'
           AND (INSTR(BD.DRUG_NAME, '奥司他韦') > 0 OR
               INSTR(BD.DRUG_NAME, '扎那米韦') > 0 OR
               INSTR(BD.DRUG_NAME, '阿比多尔') > 0 OR
               INSTR(BD.DRUG_NAME, '阿比朵尔') > 0 OR
               INSTR(BD.DRUG_NAME, '莲花清瘟') > 0 OR
               INSTR(BD.DRUG_NAME, '金花清感颗粒') > 0 OR
               INSTR(BD.DRUG_NAME, '金刚烷胺') > 0 OR
               INSTR(BD.DRUG_NAME, '金刚乙胺') > 0 OR
               INSTR(BD.DRUG_NAME, '利巴韦林') > 0))

", sTime, eTime);

            FluDataSet = GetDataSet(sqlVfp);

            return FluDataSet;

        }

        /// <summary>
        /// 检验记录数据
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button6_Click(object sender, EventArgs e)
        {
            #region 检验记录数据

            string sql = @"


--附件7 检验记录（门诊）
SELECT '01' P7501, -- 1 就诊类型 
       c.patientid P7502, --2 就诊卡号  门急诊卡号或住院号或病案号
       r.reg_date P7506, --就诊日期  
       c.barcode P8000, --4 标本号   检验标本唯一性标识
       '4' P8001, --5 流感检测代码
       d.sampletime P8002, --6 送检时间  
       c.result P8003, --7 检验结果描述  
       (CASE
         WHEN instr(c.value, '阳性') > 0 THEN
          '1'
         ELSE
          '2'
       END) P8004, --8 检验结果是否阳性 
       decode(C.Value, '阴性', '35', '阳性(+)', '36', '99') P8005 --9检测结果阳性类别 P8004=1 时必填。

  from LIS_RESULT c, lis_test_reg d, fin_opr_register r
 where c.barcode = d.barcode
   and c.patientid = r.card_no
   and d.his_itemcode = 'F00000013480'
   and c.itemcode in ('IB', 'P3', 'P2', 'P1', 'IA')
   and c.result = '阳性(+)'
   and c.patientid in (select r.card_no
                         from fin_opr_register r
                        where r.clinic_code = '{0}'
                          and r.valid_flag = '1'
                          and rownum = 1)

union all

--附件7 检验记录（住院）
SELECT '01' P7501, -- 1 就诊类型 
       c.patientid P7502, --2 就诊卡号  门急诊卡号或住院号或病案号
       r.reg_date P7506, --就诊日期  
       c.barcode P8000, --4 标本号   检验标本唯一性标识
       '4' P8001, --5 流感检测代码
       d.sampletime P8002, --6 送检时间  
       c.result P8003, --7 检验结果描述  
       (CASE
         WHEN instr(c.value, '阳性') > 0 THEN
          '1'
         ELSE
          '2'
       END) P8004, --8 检验结果是否阳性 
       decode(C.Value, '阴性', '35', '阳性(+)', '36', '99') P8005 --9检测结果阳性类别 P8004=1 时必填。

  from LIS_RESULT c, lis_test_reg d, fin_opr_register r
 where c.barcode = d.barcode
   and c.patientid = r.card_no
   and d.his_itemcode = 'F00000013480'
   and c.itemcode in ('IB', 'P3', 'P2', 'P1', 'IA')
   and c.result = '阳性(+)'
   and c.patientid in (select card_no
                         from fin_ipr_inmaininfo i
                        where i.inpatient_no = '{0}')



";

            #endregion


            #region 导出数据
            //获取流感患者
            DataSet vfp = FluDataSet;

            DataTable all = new DataTable();

            if (vfp.Tables[0].Rows.Count > 0)
            {

                DataTable dt = vfp.Tables[0];

                var query = from t in dt.AsEnumerable()
                            group t by new { t1 = t.Field<string>("clinic_code") } into m
                            select new
                            {
                                clinic_code = m.Key.t1
                            };
                if (query.ToList().Count > 0)
                {
                    query.ToList().ForEach(q =>
                    {
                        if (!q.clinic_code.Contains("-"))
                        {
                            DataSet ds = GetDataSet(string.Format(sql, q.clinic_code));

                            if (ds.Tables[0].Rows.Count > 0)
                            {
                                all.Merge(ds.Tables[0]);
                            }
                        }
                    });
                }

                if (all.Rows.Count > 0)
                {
                    string filePath = strFile + @"\lis_" + System.DateTime.Now.ToString("yyyyMMddHHmmss") + ".csv";

                    TableToCsv(all, filePath);
                    strZip = filePath.Replace("Excel", "ZipFile");
                    

                    //压缩备份
                    string strZipBak = strZip.Replace("ZipFile", "ZipFileBak").Replace(".csv", ".zip");
                    ZipFile(strFile, strZipBak);

                    File.Copy(filePath, strZip, true);
                }
            }
            #endregion

        }

        /// <summary>
        /// 死亡记录数据
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button4_Click(object sender, EventArgs e)
        {

            #region 死亡记录数据
            string sql = @"

--附件5 死亡记录接口标准 
select t.patient_no as P3, --病案号
       t.name as P4, --姓名
       decode(t.sex_code, 'F', '2', 'M', '1', '9') as P5, --性别
       to_char(sysdate, 'yyyy') - to_char(t.birthday, 'yyyy') as P7, --年龄
       t.in_date as P22, --入院日期
       (select a.new_data_name
          from com_shiftdata a
         where a.shift_type = 'B'
           and a.clinic_no = t.inpatient_no
           AND ROWNUM = 1) as P23, --入院科别
       (select a.new_data_name
          from com_shiftdata a
         where a.shift_type = 'RO'
           and a.clinic_no = t.inpatient_no
           AND ROWNUM = 1) as P24, --转科科别
       t.out_date as P25, --出院日期
       t.dept_name as P26, --出院科别,
       trunc(out_date - t.in_date) as P27, --实际住院天数
       (SELECT b.diag_name
          FROM met_cas_diagnose b
         WHERE t.inpatient_no = b.inpatient_no
           AND b.diag_kind = '11'
           AND ROWNUM = 1) as P8600, --入院诊断
       (SELECT c.value
          FROM rcd_record_item c, fin_ipr_inmaininfo t
         WHERE c.inpatient_id = t.emr_inpatientid
           AND c.value IS NOT NULL
           AND c.element_id = '754'
           AND ROWNUM = 1) as P8604, ----入院情况及诊疗和抢救经过
       (SELECT c.value
          FROM rcd_record_item c, fin_ipr_inmaininfo t
         WHERE c.inpatient_id = t.emr_inpatientid
              
           AND c.value IS NOT NULL
           AND c.element_id = '374'
           AND ROWNUM = 1) as P8605, ----死亡诊断
       nvl((select nvl(c.value, ' ')
             from pt_inpatient_cure        cure,
                  rcd_inpatient_record_set recset,
                  rcd_inpatient_record     rec,
                  rcd_record_item          c
            where cure.inpatient_code = t.inpatient_no
              and cure.id = recset.inpatient_id
              and recset.id = rec.inpatient_record_set_id
              and rec.id = c.inpatient_record
              and rec.record_child_type = 'Dead_Record'
              and c.element_id = 373
              and rownum = 1),
           (SELECT c.value
              FROM rcd_record_item c
             WHERE c.inpatient_id = t.emr_inpatientid
               AND c.value IS NOT NULL
               AND c.element_id = '373'
               AND ROWNUM = 1)) as P8606, -----死亡原因
       nvl((select to_date(substr(c.value, 0, 4) || '-' ||
                          substr(c.value,
                                 instr(c.value, '年') + 1,
                                 instr(c.value, '月') - instr(c.value, '年') - 1) || '-' ||
                          substr(c.value,
                                 instr(c.value, '月') + 1,
                                 instr(c.value, '日') - instr(c.value, '月') - 1) || ' ' ||
                          substr(c.value,
                                 instr(c.value, '日') + 1,
                                 instr(c.value, '时') - instr(c.value, '日') - 1) || ':' ||
                          substr(c.value,
                                 instr(c.value, '时') + 1,
                                 instr(c.value, '分') - instr(c.value, '时') - 1),
                          'yyyy-MM-dd HH24:mi')
             from pt_inpatient_cure        cure,
                  rcd_inpatient_record_set recset,
                  rcd_inpatient_record     rec,
                  rcd_record_item          c
            where cure.inpatient_code = t.inpatient_no
              and cure.id = recset.inpatient_id
              and recset.id = rec.inpatient_record_set_id
              and rec.id = c.inpatient_record
              and rec.record_child_type = 'Dead_Record'
              and c.element_id = 376
              and rownum = 1),
           t.out_date) as P8509 ---死亡时间 
  from fin_ipr_inmaininfo t
 WHERE t.inpatient_no = '{0}'      
   and t.zg = '4'
   AND t.in_state in ('B', 'O')


            ";
            #endregion

            #region 导出数据

            //获取流感患者
            DataSet vfp = FluDataSet;//

            DataTable all = new DataTable();

            if (vfp.Tables[0].Rows.Count > 0)
            {

                DataTable dt = vfp.Tables[0];

                DataView receiveDV = dt.DefaultView;

                receiveDV.RowFilter = "patient_type ='住院'";
                dt = receiveDV.ToTable();

                var query = from t in dt.AsEnumerable()
                            group t by new { t1 = t.Field<string>("clinic_code") } into m
                            select new
                            {
                                clinic_code = m.Key.t1
                            };
                if (query.ToList().Count > 0)
                {
                    query.ToList().ForEach(q =>
                    {

                        DataSet ds = GetDataSet(string.Format(sql, q.clinic_code));

                        if (ds.Tables[0].Rows.Count > 0)
                        {
                            all.Merge(ds.Tables[0]);
                        }
                    });
                }

                if (all.Rows.Count > 0)
                {
                    string filePath = strFile + @"\hdr_" + System.DateTime.Now.ToString("yyyyMMddHHmmss") + ".csv";

                    TableToCsv(all, filePath);
                    strZip = filePath.Replace("Excel", "ZipFile");


                    //压缩备份
                    string strZipBak = strZip.Replace("ZipFile", "ZipFileBak").Replace(".csv", ".zip");
                    ZipFile(strFile, strZipBak);

                    File.Copy(filePath, strZip, true);
                }
            }
            #endregion

        }

        private void Form1_Load(object sender, EventArgs e)
        {

            #region 创建目录

            string CreateFilePath = strFile;


            if (System.IO.Directory.Exists(CreateFilePath) == false)
            {
                System.IO.Directory.CreateDirectory(CreateFilePath);
            }

            CreateFilePath = CreateFilePath.Replace("Excel", "ZipFile");
            if (System.IO.Directory.Exists(CreateFilePath) == false)
            {
                System.IO.Directory.CreateDirectory(CreateFilePath);
            }
            CreateFilePath = CreateFilePath.Replace("ZipFile", "ZipFileBak");

            if (System.IO.Directory.Exists(CreateFilePath) == false)
            {
                System.IO.Directory.CreateDirectory(CreateFilePath);
            }
            #endregion

            try
            {
                //获取流感患者
                GetFluUserData();

                //this.label5.Text = "当前上传：门急诊和在院流感病例数据";
                //this.button1_Click(null, null);

                //this.label5.Text = "当前上传：出院流感病例数据";
                //this.button2_Click(null, null);

                //this.label5.Text = "当前上传：出院小结数据";
                //this.button3_Click(null, null);

                //this.label5.Text = "当前上传：死亡记录数据";
                //this.button4_Click(null, null);

                //this.label5.Text = "当前上传：用药记录数据";
                //this.button5_Click(null, null);

                //this.label5.Text = "当前上传：检验记录数据";
                //this.button6_Click(null, null);

                //RemoveFiles(strFile);
            }
            catch (Exception ex)
            {
                AppErrorInfor("", "", ex);

            }


            //Application.Exit();
        }

        /// <summary>
        ///  记录报错日志
        /// </summary>
        /// <param name="Name">接口名称</param>
        /// <param name="strdata">请求参数</param>
        /// <param name="e">报错信息</param>
        //读写锁，当资源处于写入模式时，其他线程写入需要等待本次写入结束之后才能继续写入
        static ReaderWriterLockSlim LogWriteLock = new ReaderWriterLockSlim();
        public static void AppErrorInfor(string Name, string strdata, Exception e)
        {
            try
            {
                //设置读写锁为写入模式独占资源，其他写入请求需要等待本次写入结束之后才能继续写入
                //注意：长时间持有读线程锁或写线程锁会使其他线程发生饥饿 (starve)。 为了得到最好的性能，需要考虑重新构造应用程序以将写访问的持续时间减少到最小。
                //      从性能方面考虑，请求进入写入模式应该紧跟文件操作之前，在此处进入写入模式仅是为了降低代码复杂度
                //      因进入与退出写入模式应在同一个try finally语句块内，所以在请求进入写入模式之前不能触发异常，否则释放次数大于请求次数将会触发异常
                LogWriteLock.EnterWriteLock();

                bool isCopy = false;

                string errorlog = AppDomain.CurrentDomain.SetupInformation.ApplicationBase + @"errorlog\";

                System.IO.DirectoryInfo dir = new System.IO.DirectoryInfo(errorlog);

                if (!System.IO.Directory.Exists(errorlog))
                {
                    System.IO.Directory.CreateDirectory(errorlog);
                }

                FileStream fs = new FileStream(AppDomain.CurrentDomain.SetupInformation.ApplicationBase + @"errorlog\error.txt", FileMode.Append);
                //超过2M
                if (fs.Length > 2097152)
                {
                    System.IO.File.Copy(AppDomain.CurrentDomain.SetupInformation.ApplicationBase + @"errorlog\error.txt", (AppDomain.CurrentDomain.SetupInformation.ApplicationBase + @"errorlog\" + @"error_" + System.DateTime.Now.ToString("yyyyMMddHHmmss") + ".txt"));
                    isCopy = true;
                }
                StreamWriter sw = new StreamWriter(fs);
                //开始写入
                sw.WriteLine("======================================" + Name + strdata + System.DateTime.Now.ToString() + "======================================\n" + e);
                //清空缓冲区
                sw.Flush();
                //关闭流
                sw.Close();
                fs.Close();
                fs.Dispose();
                sw.Dispose();

                //删除旧报错日志
                if (isCopy)
                {
                    System.IO.File.Delete(AppDomain.CurrentDomain.SetupInformation.ApplicationBase + @"errorlog\error.txt");
                }
            }
            catch (Exception ex)
            {

            }
            finally
            {
                //退出写入模式，释放资源占用
                //注意：一次请求对应一次释放
                //      若释放次数大于请求次数将会触发异常[写入锁定未经保持即被释放]
                //      若请求处理完成后未释放将会触发异常[此模式不下允许以递归方式获取写入锁定]
                LogWriteLock.ExitWriteLock();
            }
        }

        #region 出院小结获取
        public void GetEmrOutResult(string inpatientNo, out BaseObj baseObj)
        {
            baseObj = new BaseObj();

            byte[] byt = null;
            string xml = string.Empty;

            OracleConnection connection = new OracleConnection(ConnectionString);

            string sql = @"SELECT N.DATA,(SELECT DECODE(R.ZG, '3', '1', '2', '2', '1', '4', '5')
          FROM FIN_IPR_INMAININFO R
         WHERE R.INPATIENT_NO = '{0}') ZG
  FROM EMR_QCDATA M, COM_FILEINFO N
 WHERE M.INPATIENTNO = '{0}'
   AND M.ID = N.ID
   AND M.EMRNAME in ({ 1})
   AND M.STATE <> '3'";

            sql = string.Format(sql, inpatientNo, "'出院小结', '出院记录'");

            OracleCommand command = new OracleCommand(sql, connection);
            OracleDataReader Reader = command.ExecuteReader();

            while (Reader.Read())
            {
                byt = Reader[0] as byte[];
                baseObj.ExtendA = Reader[1].ToString(); //转归
            }

            if (byt != null && byt.Length > 0)
            {
                xml = System.Text.Encoding.UTF8.GetString(DecompressBytes(byt));
            }

            if (string.IsNullOrEmpty(xml))
            {
                return;
            }

            //emrMultiLineTextBox5  入院情况  memo
            //emrMultiLineTextBox6  入院诊断  id
            //emrMultiLineTextBox4  治疗经过  user01
            //emrMultiLineTextBox3  出院情况  user02
            //emrMultiLineTextBox2  出院诊断  name
            //emrMultiLineTextBox1  出院医嘱  user03
            //emrTextBox2           医生    doctname
            System.IO.File.WriteAllText("emrtmp.xml", xml);

            XElement element = XElement.Load("emrtmp.xml");

            var queryTmp = from frm in element.Elements("Object")
                           from pan in frm.Elements("Object")
                           from ele in pan.Elements("Object")
                           where ((XAttribute)ele.Attribute("name")).Value == "emrLabel7"
                           select ele;

            string title = string.Empty;

            bool isSpecial = false;

            foreach (var x in queryTmp)
            {
                title = GetPropertyValue(x);
            }

            if (title == "出 院 小 结")
            {
                isSpecial = false;
            }
            else
            {
                isSpecial = true;
            }

            if (isSpecial)
            {
                #region  产科出院小结
                var query = from frm in element.Elements("Object")
                            from pan in frm.Elements("Object")
                            from ele in pan.Elements("Object")
                            where ((XAttribute)ele.Attribute("name")).Value == "emrTextBox1" ||
                            ((XAttribute)ele.Attribute("name")).Value == "emrMultiLineTextBox3" ||
                            ((XAttribute)ele.Attribute("name")).Value == "emrTextBox20" ||
                            ((XAttribute)ele.Attribute("name")).Value == "emrTextBox17" ||
                            ((XAttribute)ele.Attribute("name")).Value == "emrTextBox14" ||
                            ((XAttribute)ele.Attribute("name")).Value == "emrDateTime2" ||
                            ((XAttribute)ele.Attribute("name")).Value == "emrComboBox8" ||
                            ((XAttribute)ele.Attribute("name")).Value == "emrComboBox7" ||
                            ((XAttribute)ele.Attribute("name")).Value == "emrTextBox7" ||
                            ((XAttribute)ele.Attribute("name")).Value == "emrComboBox6" ||
                            ((XAttribute)ele.Attribute("name")).Value == "emrTextBox6" ||
                            ((XAttribute)ele.Attribute("name")).Value == "emrTextBox5" ||
                            ((XAttribute)ele.Attribute("name")).Value == "emrTextBox4" ||
                            ((XAttribute)ele.Attribute("name")).Value == "emrComboBox5" ||
                            ((XAttribute)ele.Attribute("name")).Value == "emrComboBox4" ||
                            ((XAttribute)ele.Attribute("name")).Value == "emrTextBox3" ||
                            ((XAttribute)ele.Attribute("name")).Value == "emrComboBox2" ||
                            ((XAttribute)ele.Attribute("name")).Value == "emrTextBox2" ||
                            ((XAttribute)ele.Attribute("name")).Value == "emrComboBox1" ||
                            ((XAttribute)ele.Attribute("name")).Value == "emrComboBox3"
                            select ele;

                string emrTextBox1 = "";
                string emrMultiLineTextBox3 = "";
                string emrTextBox20 = "";
                string emrTextBox17 = "";
                string emrTextBox14 = "";
                string emrDateTime2 = "";
                string emrComboBox8 = "";
                string emrComboBox7 = "";
                string emrTextBox7 = "";
                string emrComboBox6 = "";
                string emrTextBox6 = "";
                string emrTextBox5 = "";
                string emrTextBox4 = "";
                string emrComboBox5 = "";
                string emrComboBox4 = "";
                string emrTextBox3 = "";
                string emrComboBox2 = "";
                string emrTextBox2 = "";
                string emrComboBox1 = "";
                string emrComboBox3 = "";


                foreach (var x in query)
                {
                    if (x.Attribute("name").Value == "emrTextBox1")
                    {
                        emrTextBox1 = GetPropertyValue(x);
                    }
                    if (x.Attribute("name").Value == "emrMultiLineTextBox3")
                    {
                        emrMultiLineTextBox3 = GetPropertyValue(x);
                    }
                    if (x.Attribute("name").Value == "emrTextBox20")
                    {
                        emrTextBox20 = GetPropertyValue(x);
                    }
                    if (x.Attribute("name").Value == "emrTextBox17")
                    {
                        emrTextBox17 = GetPropertyValue(x);
                    }
                    if (x.Attribute("name").Value == "emrTextBox14")
                    {
                        emrTextBox14 = GetPropertyValue(x);
                    }
                    if (x.Attribute("name").Value == "emrDateTime2")
                    {
                        emrDateTime2 = GetPropertyValue(x);
                    }
                    if (x.Attribute("name").Value == "emrComboBox8")
                    {
                        emrComboBox8 = GetPropertyValue(x);
                    }
                    if (x.Attribute("name").Value == "emrComboBox7")
                    {
                        emrComboBox7 = GetPropertyValue(x);
                    }
                    if (x.Attribute("name").Value == "emrTextBox7")
                    {
                        emrTextBox7 = GetPropertyValue(x);
                    }
                    if (x.Attribute("name").Value == "emrComboBox6")
                    {
                        emrComboBox6 = GetPropertyValue(x);
                    }
                    if (x.Attribute("name").Value == "emrTextBox6")
                    {
                        emrTextBox6 = GetPropertyValue(x);
                    }
                    if (x.Attribute("name").Value == "emrTextBox5")
                    {
                        emrTextBox5 = GetPropertyValue(x);
                    }
                    if (x.Attribute("name").Value == "emrTextBox4")
                    {
                        emrTextBox4 = GetPropertyValue(x);
                    }
                    if (x.Attribute("name").Value == "emrComboBox5")
                    {
                        emrComboBox5 = GetPropertyValue(x);
                    }
                    if (x.Attribute("name").Value == "emrComboBox4")
                    {
                        emrComboBox4 = GetPropertyValue(x);
                    }
                    if (x.Attribute("name").Value == "emrTextBox3")
                    {
                        emrTextBox3 = GetPropertyValue(x);
                    }
                    if (x.Attribute("name").Value == "emrComboBox2")
                    {
                        emrComboBox2 = GetPropertyValue(x);
                    }
                    if (x.Attribute("name").Value == "emrTextBox2")
                    {
                        emrTextBox2 = GetPropertyValue(x);
                    }
                    if (x.Attribute("name").Value == "emrComboBox1")
                    {
                        emrComboBox1 = GetPropertyValue(x);
                    }
                    if (x.Attribute("name").Value == "emrComboBox3")
                    {
                        emrComboBox3 = GetPropertyValue(x);
                    }
                }

                System.Text.StringBuilder sb = new System.Text.StringBuilder();

                sb.Append("孕").Append(emrTextBox20).Append("产").Append(emrTextBox17)
                    .Append("宫内妊娠").Append(emrTextBox14).Append("周");
                baseObj.ExtendB = sb.ToString();

                sb = new System.Text.StringBuilder();
                sb.Append("于").Append(emrDateTime2).Append(emrComboBox8).Append(emrComboBox7).Append(emrTextBox7)
                    .Append(emrComboBox6).Append("活婴。Apgar评分：1分钟").Append(emrTextBox6).Append("分，5分钟")
                    .Append(emrTextBox5).Append("分。胎盘娩出：").Append(emrTextBox4).Append("。会阴情况：裂伤")
                    .Append(emrComboBox5).Append("，切口").Append(emrComboBox4).Append("。产后出血").Append(emrTextBox3)
                    .Append("毫升（").Append(emrComboBox2).Append("），会阴伤口拆线").Append(emrTextBox2)
                    .Append("针。会阴腹部伤口愈合：").Append(emrComboBox5).Append("类，").Append(emrComboBox3);

                baseObj.ExtendC = sb.ToString();
                baseObj.ExtendD = emrTextBox1;
                baseObj.ExtendE = emrMultiLineTextBox3;

                sb = null;

                #endregion
            }
            else
            {
                #region 普通科室出院小结
                var query = from frm in element.Elements("Object")
                            from pan in frm.Elements("Object")
                            from ele in pan.Elements("Object")
                            where ((XAttribute)ele.Attribute("name")).Value == "emrMultiLineTextBox5" ||
                            ((XAttribute)ele.Attribute("name")).Value == "emrMultiLineTextBox6" ||
                            ((XAttribute)ele.Attribute("name")).Value == "emrMultiLineTextBox4" ||
                            ((XAttribute)ele.Attribute("name")).Value == "emrMultiLineTextBox3" ||
                            ((XAttribute)ele.Attribute("name")).Value == "emrMultiLineTextBox2" ||
                            ((XAttribute)ele.Attribute("name")).Value == "emrMultiLineTextBox1" ||
                            ((XAttribute)ele.Attribute("name")).Value == "emrTextBox2"
                            select ele;

                foreach (var x in query)
                {
                    if (x.Attribute("name").Value == "emrMultiLineTextBox5")
                    {
                        baseObj.ExtendB = GetPropertyValue(x);
                    }
                    //else if (x.Attribute("name").Value == "emrMultiLineTextBox6")
                    //{
                    //    baseObj.ExtendA = GetPropertyValue(x);
                    //}
                    else if (x.Attribute("name").Value == "emrMultiLineTextBox4")
                    {
                        baseObj.ExtendC = GetPropertyValue(x);
                    }
                    else if (x.Attribute("name").Value == "emrMultiLineTextBox3")
                    {
                        baseObj.ExtendD = GetPropertyValue(x);
                    }
                    //else if (x.Attribute("name").Value == "emrMultiLineTextBox2")
                    //{
                    //    neuObj.Name = GetPropertyValue(x);
                    //}
                    else if (x.Attribute("name").Value == "emrMultiLineTextBox1")
                    {
                        baseObj.ExtendE = GetPropertyValue(x);
                    }
                }
                #endregion
            }

            return;

        }

        public void GetEmrDeath(string inpatientNo, out BaseObj baseObj)
        {
            baseObj = new BaseObj();

            byte[] byt = null;
            string xml = string.Empty;

            System.Data.OracleClient.OracleConnection connect = new OracleConnection(ConnectionString);

            string sql = @"SELECT N.DATA,(SELECT DECODE(R.ZG, '3', '1', '2', '2', '1', '4', '5')
          FROM FIN_IPR_INMAININFO R
         WHERE R.INPATIENT_NO = '{0}') ZG
  FROM EMR_QCDATA M, COM_FILEINFO N
 WHERE M.INPATIENTNO = '{0}'
   AND M.ID = N.ID
   AND M.EMRNAME in ({ 1})
   AND M.STATE <> '3'";

            sql = string.Format(sql, inpatientNo, "'死亡记录'");

            OracleCommand command = new OracleCommand(sql, connect);
            OracleDataReader Reader = command.ExecuteReader();

            while (Reader.Read())
            {
                byt = Reader[0] as byte[];
                baseObj.ExtendA = Reader[1].ToString(); //转归
            }

            if (byt != null && byt.Length > 0)
            {
                xml = System.Text.Encoding.UTF8.GetString(DecompressBytes(byt));
            }

            if (string.IsNullOrEmpty(xml))
            {
                return;
            }

            //emrMultiLineTextBox5  入院情况  memo
            //emrMultiLineTextBox3  治疗经过  user01
            System.IO.File.WriteAllText("emrtmp.xml", xml);

            XElement element = XElement.Load("emrtmp.xml");

            var query = from frm in element.Elements("Object")
                        from pan in frm.Elements("Object")
                        from ele in pan.Elements("Object")
                        where ((XAttribute)ele.Attribute("name")).Value == "emrMultiLineTextBox5" ||
                        ((XAttribute)ele.Attribute("name")).Value == "emrMultiLineTextBox3"
                        select ele;

            foreach (var x in query)
            {
                if (x.Attribute("name").Value == "emrMultiLineTextBox5")
                {
                    baseObj.ExtendB = GetPropertyValue(x);
                }
                else if (x.Attribute("name").Value == "emrMultiLineTextBox3")
                {
                    baseObj.ExtendC = GetPropertyValue(x);
                }
            }

            baseObj.ExtendD = "死亡";
            baseObj.ExtendE = "无";

            return;

        }

        public void GetEmr24HourInOut(string inpatientNo, out BaseObj baseObj)
        {
            baseObj = new BaseObj();

            byte[] byt = null;
            string xml = string.Empty;

            System.Data.OracleClient.OracleConnection connect = new OracleConnection(ConnectionString);

            string sql = @"SELECT N.DATA,(SELECT DECODE(R.ZG, '3', '1', '2', '2', '1', '4', '5')
          FROM FIN_IPR_INMAININFO R
         WHERE R.INPATIENT_NO = '{0}') ZG
  FROM EMR_QCDATA M, COM_FILEINFO N
 WHERE M.INPATIENTNO = '{0}'
   AND M.ID = N.ID
   AND M.EMRNAME in ({ 1})
   AND M.STATE <> '3'";

            sql = string.Format(sql, inpatientNo, "'二十四小时内入出院记录'");

            OracleCommand command = new OracleCommand(sql, connect);
            OracleDataReader Reader = command.ExecuteReader();

            while (Reader.Read())
            {
                byt = Reader[0] as byte[];
                baseObj.ExtendA = Reader[1].ToString(); //转归
            }

            if (byt != null && byt.Length > 0)
            {
                xml = System.Text.Encoding.UTF8.GetString(DecompressBytes(byt));
            }

            if (string.IsNullOrEmpty(xml))
            {
                return;
            }

            //emrMultiLineTextBox7   入院情况  memo
            //emrMultiLineTextBox3   治疗经过  user01
            //emrMultiLineTextBox6   出院情况  user02
            //emrMultiLineTextBox5   出院医嘱  user03

            System.IO.File.WriteAllText("emrtmp.xml", xml);

            XElement element = XElement.Load("emrtmp.xml");

            var query = from frm in element.Elements("Object")
                        from pan in frm.Elements("Object")
                        from ele in pan.Elements("Object")
                        where ((XAttribute)ele.Attribute("name")).Value == "emrMultiLineTextBox5" ||
                        ((XAttribute)ele.Attribute("name")).Value == "emrMultiLineTextBox3" ||
                        ((XAttribute)ele.Attribute("name")).Value == "emrMultiLineTextBox7" ||
                        ((XAttribute)ele.Attribute("name")).Value == "emrMultiLineTextBox6"
                        select ele;

            foreach (var x in query)
            {
                if (x.Attribute("name").Value == "emrMultiLineTextBox5")
                {
                    baseObj.ExtendE = GetPropertyValue(x);
                }
                else if (x.Attribute("name").Value == "emrMultiLineTextBox3")
                {
                    baseObj.ExtendC = GetPropertyValue(x);
                }
                else if (x.Attribute("name").Value == "emrMultiLineTextBox6")
                {
                    baseObj.ExtendD = GetPropertyValue(x);
                }
                else if (x.Attribute("name").Value == "emrMultiLineTextBox7")
                {
                    baseObj.ExtendB = GetPropertyValue(x);
                }
            }

            return;

        }

        public void GetEmr24HourInOutDeath(string inpatientNo, out BaseObj baseObj)
        {
            baseObj = new BaseObj();

            byte[] byt = null;
            string xml = string.Empty;

            System.Data.OracleClient.OracleConnection connect = new OracleConnection(ConnectionString);

            string sql = @"SELECT N.DATA,(SELECT DECODE(R.ZG, '3', '1', '2', '2', '1', '4', '5')
          FROM FIN_IPR_INMAININFO R
         WHERE R.INPATIENT_NO = '{0}') ZG
  FROM EMR_QCDATA M, COM_FILEINFO N
 WHERE M.INPATIENTNO = '{0}'
   AND M.ID = N.ID
   AND M.EMRNAME in ({ 1})
   AND M.STATE <> '3'";

            sql = string.Format(sql, inpatientNo, "'二十四小时内入院死亡记录'");

            OracleCommand command = new OracleCommand(sql, connect);
            OracleDataReader Reader = command.ExecuteReader();

            while (Reader.Read())
            {
                byt = Reader[0] as byte[];
                baseObj.ExtendA = Reader[1].ToString(); //转归
            }

            if (byt != null && byt.Length > 0)
            {
                xml = System.Text.Encoding.UTF8.GetString(DecompressBytes(byt));
            }

            if (string.IsNullOrEmpty(xml))
            {
                return;
            }

            //emrMultiLineTextBox6   入院情况  memo
            //emrMultiLineTextBox4   治疗经过  user01

            System.IO.File.WriteAllText("emrtmp.xml", xml);

            XElement element = XElement.Load("emrtmp.xml");

            var query = from frm in element.Elements("Object")
                        from pan in frm.Elements("Object")
                        from ele in pan.Elements("Object")
                        where ((XAttribute)ele.Attribute("name")).Value == "emrMultiLineTextBox6" ||
                        ((XAttribute)ele.Attribute("name")).Value == "emrMultiLineTextBox4"
                        select ele;

            foreach (var x in query)
            {
                if (x.Attribute("name").Value == "emrMultiLineTextBox4")
                {
                    baseObj.ExtendC = GetPropertyValue(x);
                }
                else if (x.Attribute("name").Value == "emrMultiLineTextBox6")
                {
                    baseObj.ExtendB = GetPropertyValue(x);
                }
            }

            baseObj.ExtendD = "死亡";
            baseObj.ExtendE = "无";

            return;

        }

        public static string GetPropertyValue(XElement element)
        {
            string result = string.Empty;

            var query = from ele in element.Elements("Property")
                        where ele.Attribute("name").Value == "Value" || ele.Attribute("name").Value == "Text"
                        select ele.Value;

            foreach (var q in query)
            {
                result = q.ToString();
                break;
            }

            return result;
        }

        public static byte[] DecompressBytes(byte[] input)
        {
            using (MemoryStream mem = new MemoryStream(input))
            using (ZipInputStream stm = new ZipInputStream(mem))
            using (MemoryStream mem2 = new MemoryStream())
            {
                ZipEntry entry = stm.GetNextEntry();
                if (entry != null)
                {
                    byte[] data = new byte[4096];

                    while (true)
                    {
                        int size = stm.Read(data, 0, data.Length);
                        if (size > 0)
                        {
                            mem2.Write(data, 0, size);
                        }
                        else
                        {
                            break;
                        }
                    }
                }

                using (BinaryReader r = new BinaryReader(mem2))
                {
                    byte[] c = new byte[mem2.Length];
                    mem2.Seek(0, SeekOrigin.Begin);
                    r.Read(c, 0, (int)mem2.Length);

                    return c;
                }
            }
        }
        #endregion

    }
}
