const nodemailer = require("nodemailer");
const express = require("express");
const oracledb = require("oracledb");
const cors = require("cors");

const app = express();
const port = process.env.PORT || 7000;
const ip = process.env.PORT || "0.0.0.0";

const currentDate = new Date();
const formattedDate = currentDate.toLocaleDateString('vi-VN');

const email = {
    'hainam-nguyen@vn.apachefootwear.com' : 'S',
    'hoa-le@vn.apachefootwear.com' : 'S'
}

app.use(express.json());
app.use(cors());

//Config mail
const emailConfig = {
  host: "mail.apachefootwear.com",
  port: 25,
  secure: false,
  auth: {
    user: "APACHE\\APH-System",
    pass: "Aa123456",
  },
  tls: {
    rejectUnauthorized: false,
  },
};

//Config database
const dbConfig = {
  user: "mes00",
  password: "dbmes00",
  connectString: "10.30.3.51:1521/APHMES",
};

let clientOpts = {};
if (process.platform === "win32") {
  // Windows
  // If you use backslashes in the libDir string, you will
  // need to double them.
  clientOpts = { libDir: "C:\\oracle\\instantclient_21_11" };
} else if (process.platform === "darwin" && process.arch === "x64") {
  // macOS Intel
  clientOpts = { libDir: process.env.HOME + "/Downloads/instantclient_19_8" };
}

async function connectOracle() {
  try {
    oracledb.initOracleClient(clientOpts);
    await oracledb.createPool(dbConfig);
    console.log("Connected to Oracle");
  } catch (err) {
    console.log("Error connect to Oracle:", err);
  }
}

let emailSent = false;
connectOracle();

app.get("/api/ora", async (req, res) => {
  try {
    const connectionOracle = await oracledb.getConnection();
    const dataQueryOracle = `SELECT ROWNUM AS sno,
      d.department_name,
      d.department_code as scan_detpt,
      NVL(c.target, 0) AS target,
      NVL(c.workhours, 0) AS workhours,
      NVL(c.h07, 0) AS h07,
      NVL(c.h08, 0) AS h08,
      NVL(c.h09, 0) AS h09,
      NVL(c.h10, 0) AS h10,
      NVL(c.h11, 0) AS h11,
      NVL(c.h12, 0) AS h12,
      NVL(c.h13, 0) AS h13,
      NVL(c.h14, 0) AS h14,
      NVL(c.h15, 0) AS h15,
      NVL(c.h16, 0) AS h16,
      NVL(c.h17, 0) AS h17,
      NVL(c.h18, 0) AS h18,
      NVL(c.h19, 0) AS h19,
      NVL(c.total, 0) AS total
 FROM sjqdms_orginfo f, base005m d
 LEFT JOIN (SELECT scan_detpt,
                   pg_jms.GF_JMS_WorkDay_Target(scan_detpt, TRUNC(SYSDATE)) AS target,
                   GF_WORKINGHOURS_PERC(scan_detpt) AS workhours,
                   SUM(DECODE(hours, '07', label_qty, NULL)) AS H07,
                   SUM(DECODE(hours, '08', label_qty, NULL)) AS H08,
                   SUM(DECODE(hours, '09', label_qty, NULL)) AS H09,
                   SUM(DECODE(hours, '10', label_qty, NULL)) AS H10,
                   SUM(DECODE(hours, '11', label_qty, NULL)) AS H11,
                   SUM(DECODE(hours, '12', label_qty, NULL)) AS H12,
                   SUM(DECODE(hours, '13', label_qty, NULL)) AS H13,
                   SUM(DECODE(hours, '14', label_qty, NULL)) AS H14,
                   SUM(DECODE(hours, '15', label_qty, NULL)) AS H15,
                   SUM(DECODE(hours, '16', label_qty, NULL)) AS H16,
                   SUM(DECODE(hours, '17', label_qty, NULL)) AS H17,
                   SUM(DECODE(hours, '18', label_qty, NULL)) AS H18,
                   SUM(DECODE(hours, '19', label_qty, NULL)) AS H19,
                   SUM(label_qty) AS Total,
                   SUM(label_qty) -
                   pg_jms.GF_JMS_WorkDay_Target(scan_detpt, TRUNC(SYSDATE)) AS Balance,
                   CASE
                     WHEN NVL(MAX(pg_jms.GF_JMS_WorkDay_Target(scan_detpt,
                                                               TRUNC(SYSDATE))),
                              0) = 0 THEN
                      '无目标产量'
                     WHEN GF_WORKINGHOURS_PERC(scan_detpt) = 0 THEN
                      '无工作历资料'
                     ELSE
                      NVL(ROUND(SUM(label_qty) /
                                MAX(pg_jms.GF_JMS_WorkDay_Target(scan_detpt,
                                                                 TRUNC(SYSDATE))),
                                3) * 100,
                          0) || '%'
                   END AS Acheivement,
                   CASE
                     WHEN NVL(MAX(pg_jms.GF_JMS_WorkDay_Target(scan_detpt,
                                                               TRUNC(SYSDATE))),
                              0) = 0 THEN
                      0
                     ELSE
                      NVL(ROUND(SUM(label_qty) /
                                MAX(pg_jms.GF_JMS_WorkDay_Target(scan_detpt,
                                                                 TRUNC(SYSDATE))),
                                3),
                          0) * 100
                   END AS Bcheivement,
                   GF_WORKINGHOURS_PERC(scan_detpt) AS Ccheivement
              FROM (SELECT SUM(label_qty) AS label_qty, hours, scan_detpt
                      FROM (SELECT scan_detpt,
                                   scan_date,
                                   TO_CHAR((TO_DATE(scan_date,
                                                    'yyyy/mm/dd hh24:mi:ss') -
                                           30 / 24 / 60),
                                           'hh24') AS hours,
                                   SUM(label_qty) AS label_qty
                              FROM (SELECT SUM(label_qty) AS label_qty,
                                           TO_CHAR(scan_date, 'mi'),
                                           nvl(scan_detpt, 0) as scan_detpt,
                                           (CASE
                                             WHEN FLOOR((TO_CHAR(scan_date,
                                                                 'mi')) / 30) = 0 THEN
                                              TO_CHAR(scan_date,
                                                      'yyyy/mm/dd hh24') ||
                                              ':00:00'
                                             WHEN FLOOR((TO_CHAR(scan_date,
                                                                 'mi')) / 30) = 1 THEN
                                              TO_CHAR(scan_date,
                                                      'yyyy/mm/dd hh24') ||
                                              ':30:00'
                                           END) AS scan_date
                                      FROM sfc_trackout_list s
                                     WHERE scan_date >= TRUNC(SYSDATE)
                                       AND scan_date < TRUNC(SYSDATE + 1)
                                       AND INOUT_PZ = 'OUT'
                                     GROUP BY scan_detpt, scan_date)
                             GROUP BY scan_detpt, scan_date)
                     GROUP BY hours, scan_detpt
                     ORDER BY scan_detpt, hours)
             GROUP BY scan_detpt) c
   ON d.department_code = c.scan_detpt
WHERE f.code = d.udf05
  AND d.udf06 = 'Y'
  AND (d.department_code like '%S%')
ORDER BY d.department_code asc`;
    const dataResultOracle = await connectionOracle.execute(dataQueryOracle);
    const dataOracle = dataResultOracle.rows;
    connectionOracle.close();

    const jsonDataOracle = dataOracle.map((row) => ({
      //DEPARTMENT_CODE: row[0],
      DEPARTMENT_NAME: row[1],
      SCAN_DEPT: row[2],
      H07: row[5],
      H08: row[6],
      H09: row[7],
      H010: row[8],
      H11: row[9],
      H12: row[10],
      H13: row[11],
      H14: row[12],
      H15: row[13],
      H16: row[14],
      H17: row[15],
      H18: row[16],
      H19: row[17],
      TOTAL: row[18],
    }));

    const htmlReport = HTMLReport(dataOracle)


    const data = Object.entries(email).map(([key, value]) => ({
        key,
        value,
      }));
      // Send emails
      data.forEach((item) => {
        sendMail(item.key, item.value);
      });
    
      res.send("Emails sent successfully");
    res.json(htmlReport);

    await sendMail(htmlReport);
  } catch (err) {
    console.error("Error executing Oracle query:", err);
    res.status(500).send("Error fetching data from Oracle database");
  }
});

async function sendMail() {
  try {
    const transporter = nodemailer.createTransport(emailConfig);

    const toMail = "hainam-nguyen@vn.apachefootwear.com";
    const htmlBody = generateHtmlBody(item);


    const mailOptions = {
        from: '"APH-System" <APH-System@vn.apachefootwear.com>',
        to: plants,
        subject: 'Stitching Report',
        html: htmlBody,
      };
    
      try {
        const info = await transporter.sendMail(mailOptions);
        console.log("Email sent to:", info.messageId);
      } catch (error) {
        console.error("Error sending email:", error);
      }

    //Log mail address
    console.log("Send mail to:", toMail);

    const info = await transporter.sendMail(mailOptions);
    console.log("Email send to:", info.messageId);
  } catch (error) {
    console.log("Error send Mail: ", error);
  }
}

function HTMLReport(item){
    let textBody = `
    <br>
    <style>
        table {
          font-family: '等线';
          font-size: 15px;
          border-collapse: collapse;
        }
        td, th {
          border: 1px solid black;
          text-align: left;
          padding: 4px;
          color: #003366;
        }
        th {
          background-color: #ADD8E6;
        }
        th {
          color: black;
        }
        td {
          background-color: #FFFFE0;
        }
        .border {
          background-color: transparent;
          border: none;
        }
    </style>
    <h3>${formattedDate}</h3>
    <table role="presentation">
        <tr>
            <td class="border" align="center">
                <table role="presentation">
                    <tr>
                        <th style="width: 40%">Index<br>(序号)</th>
                        <th style="width: 30%">Plant<br>(車間)</th>
                        <th style="width: 30%">7:30 - 8:30</th>
                        <th style="width: 30%">8:30 - 9:30</th>
                        <th style="width: 30%">9:30 - 10:30</th>
                        <th style="width: 30%">11:30 - 12:30</th>
                        <th style="width: 30%">12:30 - 13:30</th>
                        <th style="width: 30%">13:30 - 14:30</th>
                        <th style="width: 30%">14:30 - 15:30</th>
                        <th style="width: 30%">15:30 - 16:30</th>
                        <th style="width: 30%">16:30 - 17:30</th>
                        <th style="width: 30%">17:30 - 18:30</th>
                        <th style="width: 30%">18:30 - 19:30</th>
                        <th style="width: 30%">19:30 - 20:30</th>
                    </tr>`;

  // Example data (replace this with your actual data)
  const table = dic[item];

  for (let i = 0; i < table.length; i++) {
    textBody += `
                    <tr>
                        <td>${i + 1}</td>
                        <td>${table[i][1]}</td>
                        <td>${table[i][5]}</td>
                        <td>${table[i][6]}</td>
                        <td>${table[i][7]}</td>
                        <td>${table[i][8]}</td>
                        <td>${table[i][9]}</td>
                        <td>${table[i][10]}</td>
                        <td>${table[i][11]}</td>
                        <td>${table[i][12]}</td>
                        <td>${table[i][13]}</td>
                        <td>${table[i][14]}</td>
                        <td>${table[i][15]}</td>
                        <td>${table[i][16]}</td>
                    </tr>`;
  }
  textBody += `
                </table>
            </td>
        </tr>
    </table>`;

  return textBody;
}
app.listen(port, ip, () => {
  console.log(`Server listening on ${ip}:${port}`);
});

setInterval(checkVolumeAndsendEmail, 19000);
