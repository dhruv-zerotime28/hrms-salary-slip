import {
  Document,
  Page,
  View,
  Text,
  Image,
  StyleSheet,
} from "@react-pdf/renderer";

export type EmployeeData = {
  name: string;
  panNumber: string;
  designation: string;
  doj: string;
  employeeId: string;
  totalDays: string;
  payableDays: string;
  absentDays: string;
  basicDA: string;
  hra: string;
  lta: string;
  allowance: string;
  grossSalary: string;
  other: string;
  retentionBonus: string;
  pt: string;
  tds: string;
  pf: string;
  esic: string;
  earningTotal: string;
  deductionTotal: string;
  netPay: string;
  netPayInWords: string;
  bankAc: string;
  ifsc: string;
  branch: string;
};

const s = StyleSheet.create({
  page: {
    padding: "14mm",
    paddingBottom: "40mm",
    fontSize: 10,
    fontFamily: "Helvetica",
  },
  topLine: { borderTopWidth: 3, borderTopColor: "#000", marginBottom: 10 },
  bottomLine: {
    borderBottomWidth: 3,
    borderBottomColor: "#000",
    position: "absolute",
    bottom: "10mm",
    left: "14mm",
    right: "14mm",
  },
  // Header
  logo: { width: 140, height: 60, objectFit: "contain" as const },
  headerCenter: { textAlign: "center", marginTop: 10, marginBottom: 2 },
  companyName: {
    fontSize: 19,
    fontWeight: "bold",
    color: "#000",
    marginBottom: 6,
    fontFamily: "Helvetica-Bold",
  },
  slipTitle: {
    fontSize: 15,
    fontWeight: "bold",
    color: "#000",
    textDecoration: "underline",
    fontFamily: "Helvetica-Bold",
  },
  // Table shared
  table: {
    marginTop: 14,
    borderTopWidth: 1,
    borderTopColor: "#000",
    borderLeftWidth: 1,
    borderLeftColor: "#000",
  },
  row: { flexDirection: "row" },
  // Info table
  labelCell: {
    width: "22%",
    borderRightWidth: 1,
    borderRightColor: "#000",
    borderBottomWidth: 1,
    borderBottomColor: "#000",
    padding: "5 8",
    fontFamily: "Helvetica-Bold",
  },
  valueCell: {
    width: "28%",
    borderRightWidth: 1,
    borderRightColor: "#000",
    borderBottomWidth: 1,
    borderBottomColor: "#000",
    padding: "5 8",
  },
  // Salary table
  salaryCell30: {
    width: "30%",
    borderRightWidth: 1,
    borderRightColor: "#000",
    borderBottomWidth: 1,
    borderBottomColor: "#000",
    padding: "5 8",
  },
  salaryCell20: {
    width: "20%",
    borderRightWidth: 1,
    borderRightColor: "#000",
    borderBottomWidth: 1,
    borderBottomColor: "#000",
    padding: "5 8",
  },
  headerCell30: {
    width: "30%",
    borderRightWidth: 1,
    borderRightColor: "#000",
    borderBottomWidth: 1,
    borderBottomColor: "#000",
    padding: "6 8",
    fontFamily: "Helvetica-Bold",
  },
  headerCell20: {
    width: "20%",
    borderRightWidth: 1,
    borderRightColor: "#000",
    borderBottomWidth: 1,
    borderBottomColor: "#000",
    padding: "6 8",
    fontFamily: "Helvetica-Bold",
  },
  bold: { fontFamily: "Helvetica-Bold" },
  // Net pay
  netPayLabel: {
    width: "40%",
    borderRightWidth: 1,
    borderRightColor: "#000",
    borderBottomWidth: 1,
    borderBottomColor: "#000",
    padding: "6 8",
    fontFamily: "Helvetica-Bold",
  },
  netPayValue: {
    width: "60%",
    borderRightWidth: 1,
    borderRightColor: "#000",
    borderBottomWidth: 1,
    borderBottomColor: "#000",
    padding: "6 8",
    fontFamily: "Helvetica-Bold",
  },
  // Footer
 footer: {
    position: "absolute",
    bottom: "12mm",
    left: "14mm",
    right: "14mm",
    fontSize: 9,
    textAlign: "right",
  },
  footerLine: { marginBottom: 3 },
});

function formatINR(value: string | number | undefined): string {
  if (value === undefined || value === "" || value === "-") return "-";
  const num = typeof value === "string" ? parseFloat(value) : value;
  if (isNaN(num)) return String(value);
  const fixed = Math.round(num).toString();
  const isNeg = fixed.startsWith("-");
  const abs = isNeg ? fixed.slice(1) : fixed;
  if (abs.length <= 3) return (isNeg ? "-" : "") + abs;
  const last3 = abs.slice(-3);
  const rest = abs.slice(0, -3);
  const formatted = rest.replace(/\B(?=(\d{2})+(?!\d))/g, ",") + "," + last3;
  return (isNeg ? "-" : "") + formatted;
}

function InfoRow({
  l1,
  v1,
  l2,
  v2,
}: {
  l1: string;
  v1: string;
  l2: string;
  v2: string;
}) {
  return (
    <View style={s.row}>
      <View style={s.labelCell}>
        <Text>{l1}</Text>
      </View>
      <View style={s.valueCell}>
        <Text>{v1}</Text>
      </View>
      <View style={s.labelCell}>
        <Text>{l2}</Text>
      </View>
      <View style={s.valueCell}>
        <Text>{v2}</Text>
      </View>
    </View>
  );
}

function SalaryRow({
  earning,
  eAmt,
  deduction,
  dAmt,
  bold: isBold,
}: {
  earning: string;
  eAmt: string;
  deduction: string;
  dAmt: string;
  bold?: boolean;
}) {
  const t = isBold ? s.bold : undefined;
  return (
    <View style={s.row}>
      <View style={s.salaryCell30}>
        <Text style={t}>{earning}</Text>
      </View>
      <View style={s.salaryCell20}>
        <Text style={t}>{eAmt}</Text>
      </View>
      <View style={s.salaryCell30}>
        <Text style={t}>{deduction}</Text>
      </View>
      <View style={s.salaryCell20}>
        <Text style={t}>{dAmt}</Text>
      </View>
    </View>
  );
}

export function SalarySlipDocument({
  data,
  month,
  year,
  logoSrc,
}: {
  data: EmployeeData;
  month: string;
  year: string;
  logoSrc: string;
}) {
  return (
    <Document>
      <Page size="A4" style={s.page}>
        {/* Top border */}
        <View style={s.topLine} />

        {/* Header */}
        <Image src={logoSrc} style={s.logo} />

        <View style={s.headerCenter}>
          <Text style={s.companyName}>Zerotime Solutions {"  "} LLP</Text>
          <Text style={s.slipTitle}>
            Salary Slip for {month} {year}
          </Text>
        </View>

        {/* Employee Info Table */}
        <View style={s.table}>
          <InfoRow
            l1="Name"
            v1={data.name}
            l2="Total Days"
            v2={data.totalDays}
          />
          <InfoRow
            l1="Designation"
            v1={data.designation}
            l2="Pay Days"
            v2={data.payableDays}
          />
          <InfoRow
            l1="Joining Date"
            v1={data.doj}
            l2="Absent Days"
            v2={data.absentDays}
          />
          <InfoRow
            l1="PAN Number"
            v1={data.panNumber}
            l2="Employee ID"
            v2={data.employeeId}
          />
          <InfoRow l1="A/C Number" v1={data.bankAc} l2="IFSC" v2={data.ifsc} />
        </View>

        {/* Earnings & Deductions Table */}
        <View style={s.table}>
          <View style={s.row}>
            <View style={s.headerCell30}>
              <Text>Earning</Text>
            </View>
            <View style={s.headerCell20}>
              <Text>Amount</Text>
            </View>
            <View style={s.headerCell30}>
              <Text>Deductions</Text>
            </View>
            <View style={s.headerCell20}>
              <Text>Amount</Text>
            </View>
          </View>
          <SalaryRow
            earning="Basic + DA"
            eAmt={formatINR(data.basicDA)}
            deduction="Retention Bonus"
            dAmt={formatINR(data.retentionBonus)}
          />
          <SalaryRow
            earning="HRA"
            eAmt={formatINR(data.hra)}
            deduction="PT"
            dAmt={formatINR(data.pt)}
          />
          <SalaryRow
            earning="LTA"
            eAmt={formatINR(data.lta)}
            deduction="TDS"
            dAmt={formatINR(data.tds)}
          />
          <SalaryRow
            earning="Allowances"
            eAmt={formatINR(data.allowance)}
            deduction="PF"
            dAmt={formatINR(data.pf)}
          />
          <SalaryRow
            earning="Gross Salary"
            eAmt={formatINR(data.grossSalary)}
            deduction="ESIC"
            dAmt={formatINR(data.esic)}
          />
          <SalaryRow
            earning="Other (OT/Reb.)"
            eAmt={formatINR(data.other)}
            deduction=""
            dAmt=""
          />
          <SalaryRow
            earning="Total"
            eAmt={formatINR(data.earningTotal)}
            deduction="Total"
            dAmt={formatINR(data.deductionTotal)}
            bold
          />
        </View>

        {/* Net Pay Table */}
        <View style={s.table}>
          <View style={s.row}>
            <View style={s.netPayLabel}>
              <Text>Net Pay</Text>
            </View>
            <View style={s.netPayValue}>
              <Text>Rs. {formatINR(data.netPay)} /-</Text>
            </View>
          </View>
          <View style={s.row}>
            <View style={s.netPayLabel}>
              <Text>In Words</Text>
            </View>
            <View style={s.netPayValue}>
              <Text>{data.netPayInWords}</Text>
            </View>
          </View>
        </View>

        {/* Footer */}
        <View style={s.footer}>
          <Text style={s.footerLine}>Contact: 9409094181</Text>
          <Text style={s.footerLine}>
            232, Iscon Emporio, Near Star Bazaar, Jodhpur Cross Road, Satellite
          </Text>
          <Text style={s.footerLine}>hr@zerotimesolutions.com</Text>
          <Text style={s.footerLine}>zerotimesolutions.com</Text>
        </View>

        {/* Bottom border */}
        <View style={s.bottomLine} />
      </Page>
    </Document>
  );
}
