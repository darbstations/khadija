import {
  BarChart,
  Bar as RBar,
  XAxis,
  YAxis,
  Tooltip,
  ResponsiveContainer,
  Legend,
  CartesianGrid,
} from "recharts";
import { useScenario } from "../context/ScenarioContext";
import { projections, fmtInt, fmtNum } from "../model/engine";
import { Card, Stat } from "./ui";

export default function Projections() {
  const { inputs } = useScenario();
  const p = projections(inputs);

  const data = p.years.map((y, idx) => ({
    year: String(y),
    تشغيل: +p.operating[idx].toFixed(3),
    استثمار: +p.investment[idx].toFixed(3),
    امتياز: +p.franchise[idx].toFixed(3),
    محطات: p.totalStations[idx],
  }));

  return (
    <div className="space-y-4">
      <div>
        <h2 className="text-lg font-extrabold">📈 التوقعات السنوية (5 سنوات)</h2>
        <p className="text-xs text-darb-mut">
          التكلفة على درب بنمو المحطات من {fmtInt(p.totalStations[0])} إلى{" "}
          {fmtInt(p.totalStations[4])}
        </p>
      </div>

      <div className="grid sm:grid-cols-3 gap-3">
        <Stat label="💎 إجمالي 5 سنوات على درب" value={`${fmtNum(p.fiveYearTotal, 2)} م﷼`} tone="warn" />
        <Stat label="تكلفة 2026" value={`${fmtNum(p.total[0], 2)} م﷼`} />
        <Stat label="تكلفة 2030" value={`${fmtNum(p.total[4], 2)} م﷼`} />
      </div>

      <Card title="💰 التكلفة السنوية على درب (مليون ريال)">
        <div style={{ width: "100%", height: 320 }}>
          <ResponsiveContainer>
            <BarChart data={data} margin={{ top: 8, right: 8, left: 8, bottom: 8 }}>
              <CartesianGrid strokeDasharray="3 3" stroke="#3b3d41" />
              <XAxis dataKey="year" tick={{ fill: "#9a9da2", fontSize: 12 }} />
              <YAxis tick={{ fill: "#9a9da2", fontSize: 12 }} />
              <Tooltip
                contentStyle={{
                  background: "#2b2c2f",
                  border: "1px solid #3b3d41",
                  borderRadius: 12,
                  color: "#f4f5f6",
                }}
              />
              <Legend wrapperStyle={{ fontSize: 12 }} />
              <RBar dataKey="تشغيل" stackId="a" fill="#F18A2B" radius={[0, 0, 0, 0]} />
              <RBar dataKey="استثمار" stackId="a" fill="#9a9da2" />
              <RBar dataKey="امتياز" stackId="a" fill="#6D6E70" radius={[6, 6, 0, 0]} />
            </BarChart>
          </ResponsiveContainer>
        </div>
      </Card>

      <Card title="🏗️ تفاصيل النمو والتكلفة">
        <div className="overflow-x-auto">
          <table className="w-full border-collapse">
            <thead>
              <tr>
                <th className="th">البند</th>
                {p.years.map((y) => (
                  <th key={y} className="th">
                    {y}
                  </th>
                ))}
              </tr>
            </thead>
            <tbody>
              <tr>
                <td className="td font-bold">📊 إجمالي المحطات</td>
                {p.totalStations.map((v, idx) => (
                  <td key={idx} className="td">
                    {fmtInt(v)}
                  </td>
                ))}
              </tr>
              <tr>
                <td className="td font-bold">تكلفة التشغيل (م﷼)</td>
                {p.operating.map((v, idx) => (
                  <td key={idx} className="td">
                    {fmtNum(v, 2)}
                  </td>
                ))}
              </tr>
              <tr>
                <td className="td font-bold">تكلفة الاستثمار (م﷼)</td>
                {p.investment.map((v, idx) => (
                  <td key={idx} className="td">
                    {fmtNum(v, 2)}
                  </td>
                ))}
              </tr>
              <tr>
                <td className="td font-bold">تكلفة الامتياز (م﷼)</td>
                {p.franchise.map((v, idx) => (
                  <td key={idx} className="td">
                    {fmtNum(v, 2)}
                  </td>
                ))}
              </tr>
              <tr className="bg-darb-panel/60 font-bold">
                <td className="td">📊 الإجمالي (م﷼)</td>
                {p.total.map((v, idx) => (
                  <td key={idx} className="td text-darb-warn">
                    {fmtNum(v, 2)}
                  </td>
                ))}
              </tr>
            </tbody>
          </table>
        </div>
      </Card>
    </div>
  );
}
