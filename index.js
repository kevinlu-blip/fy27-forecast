import dynamic from "next/dynamic";

const FY27Forecast = dynamic(
  () => import("../components/FY27_ARR_Forecast"),
  { ssr: false }
);

export default function Home() {
  return <FY27Forecast />;
}
