<template>
	<div>
		<spreadsheet ref="spreadsheet" :toolbar="false" :toolbarData="false" :toolbarInsert="false" :sheetsbar="true"
			:columns="27">
			<spreadsheet-sheet :name="'MAWP'" :data-source="datasource" :rows="rows" :columns="columns" :frozenRows="1"
				:excel-proxy-URL="'https://demos.telerik.com/kendo-ui/service/export'"
				:pdf-proxy-URL="'https://demos.telerik.com/kendo-ui/service/export'">
			</spreadsheet-sheet>

			<spreadsheet-sheet :name="'BOP'" :data-source="datasource" :rows="rows" :columns="columns" :frozenRows="1"
				:excel-proxy-URL="'https://demos.telerik.com/kendo-ui/service/export'"
				:pdf-proxy-URL="'https://demos.telerik.com/kendo-ui/service/export'">
			</spreadsheet-sheet>

		</spreadsheet>
	</div>
</template>

<script>
	import {
  Spreadsheet,
  SpreadsheetSheet
} from "@progress/kendo-spreadsheet-vue-wrapper";

window.JSZip = JSZip;

export default {
  name: "App",
  components: {
    spreadsheet: Spreadsheet,
    "spreadsheet-sheet": SpreadsheetSheet
  },
  methods: {
    calc: function() {
      var rangeLength = this.datalen + 1;

      var spreadsheet = this.$refs.spreadsheet.kendoWidget();
      var sheet = spreadsheet.activeSheet();

      var rgeDeepestExposedShoeDepth = "E3:E" + rangeLength;
      var rgeG2P = "J2:J" + rangeLength;
      var rgeColumnLength = "K2:K" + rangeLength;
      var rgeColumnPercent = "L2:L" + rangeLength;
      var rgeGasColumnGradient = "M2:M" + rangeLength;
      var rgeWellheadDepth = "N2:N" + rangeLength;
      var rgeGasColumnLength = "O2:O" + rangeLength;
      var rgeGasColumnPressure = "P2:P" + rangeLength;
      var rgeMudColumnPressure = "Q2:Q" + rangeLength;
      var rgeFracpressureatshoe = "R2:R" + rangeLength;
      var rgeFracPressureAtWellheadIfAllGas = "S2:S" + rangeLength;
      var rgeFracPressureAtWellheadIfPartiallyFluid = "T2:T" + rangeLength;
      var rgeNoFracPorePressureOnBottom = "U2:U" + rangeLength;
      var rgePorePartialGasASP = "V2:V" + rangeLength;
      var rgeBHPatdeepestexposeddepth = "W2:W" + rangeLength;
      var rgeOneThirdBHPatsurface = "X2:X" + rangeLength;
      var rgeMAWPBSEE = "Y2:Y" + rangeLength;
      var rgeMAWPShell = "Z2:Z" + rangeLength;
      var rgeMAWP = "AA2:AA" + rangeLength;

      sheet.range("E2").formula("F2");
      sheet.range(rgeDeepestExposedShoeDepth).formula("F2");

      sheet.range(rgeG2P).value(0.052);
      sheet.range(rgeG2P).enable(false);

      sheet.range(rgeColumnLength).formula("F2-(C2+D2)");
      sheet
        .range(rgeColumnPercent)
        .formula(
          "IF(K2<=10000,0.7,IF(AND(K2>10001,K2<12000),0.6-(0.1/2000)*(K2-12000),IF(K2=12000,0.6,IF(AND(K2>12000,K2<15000),0.5-(0.1/3000)*(K2-15000),IF(K2>=15000,0.5)))))"
        );

      sheet
        .range(rgeGasColumnGradient)
        .formula(
          "IF(K2<=9000,0.1,IF(AND(K2>9000,K2<11000),0.15+(0.05/2000)*(K2-11000),IF(K2>=11000,0.15)))"
        );

      sheet.range(rgeWellheadDepth).formula("C2+D2");
      sheet.range(rgeGasColumnLength).formula("K2*L2");
      sheet.range(rgeGasColumnPressure).formula("K2*M2");

      sheet.range(rgeMudColumnPressure).formula("K2*(1-L2)*H2*J2");
      sheet.range(rgeFracpressureatshoe).formula("E2*I2*J2");
      sheet.range(rgeFracPressureAtWellheadIfAllGas).formula("R2-P2");
      sheet
        .range(rgeFracPressureAtWellheadIfPartiallyFluid)
        .formula("S2-(E2-N2-O2)*H2*J2");
      sheet.range(rgeNoFracPorePressureOnBottom).formula("I2*J2*F2");
      sheet.range(rgePorePartialGasASP).formula("U2-Q2-P2");

      sheet.range(rgeBHPatdeepestexposeddepth).formula("H2*J2*F2");
      sheet.range(rgeOneThirdBHPatsurface).formula("W2/3");
      sheet
        .range(rgeMAWPBSEE)
        .formula(
          "IF(S2>T2,IF(S2>V2,V2,IF(S2<=V2,S2)),IF(S2<=T2,IF(T2>V2,V2,IF(T2<=V2,T2))))"
        );
        //IF(IF(S3>T3, S3, T3)>V3, V3, IF(S3>T3, S3, T3))
      sheet.range(rgeMAWPShell).formula("(X2+(R2-X2)/E2)*N2");
      sheet.range(rgeMAWP).formula("IF(Y2>Z2,Z2,IF(Y2<=Z2,Z2))");
      //IF(Y3>Z3, Y3, Z3) 

      //Format Header Cells
      var headerCellsRange = "A1:AA1";
      sheet
        .range(headerCellsRange)
        .enable(false).borderBottom({ size: 1, color: "darkgrey" }).borderRight({ size: 1, color: "darkgrey" })
        .background("#FCF3CF")
        .bold(true)
        .textAlign("center")
        .color("black").wrap(true);
    }
  },
  mounted: function() {
    var spreadsheet = this.$refs.spreadsheet.kendoWidget();
    spreadsheet.element.css("height", "600px");
    spreadsheet.element.css("width", "100%");
    spreadsheet.resize();

    var sheet = spreadsheet.activeSheet();
    sheet.range("A1").input("Hole Size");
    sheet.range("B1").input("Nominal Casing Size");
    sheet.range("C1").input("Water Depth");
    sheet.range("D1").input("Elevation KB");
    sheet.range("E1").input("Deepest Exposed Shoe Depth");
    sheet.range("F1").input("Deepest Exposed Section Depth");
    sheet.range("G1").input("Frac Gradient");
    sheet.range("H1").input("Mud Weight");
    sheet.range("I1").input("Pore Pressure");

    sheet.range("J1").input("G2P (conversion of mud gradient to psi)");
    sheet.range("K1").input("Column Length ");
    sheet.range("L1").input("Gas Column % (ft)");
    sheet.range("M1").input("Gas Column Gradient (psi/ft)");
    sheet.range("N1").input("Wellhead Depth ");
    sheet.range("O1").input("Gas Column Length");
    sheet.range("P1").input("Gas Column Pressure");
    sheet.range("Q1").input("Mud Column Pressure");
    sheet.range("R1").input("Frac pressure at Shoe");
    sheet.range("S1").input("Frac Pressure At Wellhead (If All Gas)");
    sheet.range("T1").input("Frac Pressure At Wellhead (If Partially Fluid)");
    sheet.range("U1").input("No Frac Pore Pressure On Bottom");
    sheet.range("V1").input("Pore + Partial Gas ASP ");
    sheet.range("W1").input("BHP at Deepest Exposed Depth");
    sheet.range("X1").input("1/3 * BHP at Surface");
    sheet.range("Y1").input("MAWP (BSEE)");
    sheet.range("Z1").input("MAWP (Shell)");
    sheet.range("AA1").input("MAWP");

    this.datalen= this.datasource.data.length
    var rangeLength = this.datalen + 1;
    var dataCells="E2:E" + rangeLength;
    sheet.range(dataCells).background("#F0F3F4");
    dataCells="K2:AA" + rangeLength;
    sheet.range(dataCells).background("#F0F3F4");

    var MAWPCells="X2:AA" + rangeLength;

    sheet.range(MAWPCells).format(kendo.spreadsheet.formats.number);
    dataCells="A2:AA" + rangeLength;
    sheet.range(dataCells).borderBottom({ size: 1, color: "darkgrey" }).borderRight({ size: 1, color: "darkgrey" });

    this.calc();
  },

  data: function() {
    return {
      datalen: 0,
      rows: [
        {
          height: 40
        }
      ],
      columns: [
        { autoWidth: true },
        { width: 145 },
        { width: 145 },
        { width: 145 },
        { width: 145 },
        { width: 145 },
        { width: 145 },
        { width: 145 },
        { width: 145 },
        { width: 145 },
        { width: 145 },
        { width: 145 },
        { width: 145 },
        { width: 145 },
        { width: 145 },
        { width: 145 },
        { width: 145 },
        { width: 145 },
        { width: 145 },
        { width: 145 },
        { width: 145 },
        { width: 145 },
        { width: 145 },
        { width: 145 },
        { width: 145 },
        { width: 145 },
        { width: 145 }
      ],
      datasource: {
        data: [
          {
            HoleSize: "60",
            NominalCasingSize: "58",
            WaterDepth: "100",
            ElevationKB: "1000",
            DeepestExposedShoeDepth: "",
            DeepestExposedSectionDepth: "4000",
            FracGradient: "10",
            MudWeight: "11",
            PorePressure: "12"
          },
          {
            HoleSize: "50",
            NominalCasingSize: "48",
            WaterDepth: "100",
            ElevationKB: "1000",
            DeepestExposedShoeDepth: "",
            DeepestExposedSectionDepth: "5000",
            FracGradient: "11",
            MudWeight: "12",
            PorePressure: "13"
          },
          {
            HoleSize: "40",
            NominalCasingSize: "38",
            WaterDepth: "100",
            ElevationKB: "1000",
            DeepestExposedShoeDepth: "",
            DeepestExposedSectionDepth: "6000",
            FracGradient: "12",
            MudWeight: "13",
            PorePressure: "14"
          },
          {
            HoleSize: "30",
            NominalCasingSize: "28",
            WaterDepth: "100",
            ElevationKB: "1000",
            DeepestExposedShoeDepth: "",
            DeepestExposedSectionDepth: "7000",
            FracGradient: "13",
            MudWeight: "14",
            PorePressure: "15"
          }
        ],
        batch: true,
        schema: {
          model: {
            fields: {
              HoleSize: { type: "number", editable: false, title: "Hole Size" },
              NominalCasingSize: { type: "number", editable: false },
              WaterDepth: { type: "number", editable: false },
              ElevationKB: { type: "number", editable: false },
              DeepestExposedShoeDepth: { type: "number", editable: false },
              DeepestExposedSectionDepth: { type: "number", editable: false },
              FracGradient: { type: "number", editable: false },
              MudWeight: { type: "number", editable: false },
              PorePressure: { type: "number", editable: false }
            }
          }
        }
      }
    };
  }
};
</script>