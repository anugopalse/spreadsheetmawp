<template>
	<div>
		<!--	<spreadsheet ref="spreadsheet" :toolbarData="false" :toolbarInsert="false" :sheetsbar="false" :columns="20"
			:rows="100" :sheetsbar="false"> -->

		<spreadsheet ref="spreadsheet" :toolbar="false" :toolbarData="false" :toolbarInsert="false" :sheetsbar="false"
			:columns="16">
			<spreadsheet-sheet :name="'Products'" :data-source="datasource" :rows="rows" :columns="columns"
				:frozenRows="1" :excel-proxy-URL="'https://demos.telerik.com/kendo-ui/service/export'"
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
  methods:{
    requestEnd: function(e) {/*
      var length = e.response.length;
      var spread = this.$refs.spreadsheet.kendoWidget();

      setTimeout(function(){
        for(let i=2; i<length; i++){
            var sheet = spread.activeSheet();
            var range = sheet.range("F"+i).formula("C" +i+"*D"+i);
        }
      })*/
      this.datalen=e.response.length;
      this.calc();
    },
    calc: function(){
      // var spreadsheet = this.$refs.spreadsheet.kendoWidget();
      // spreadsheet.element.css("height", "600px");
      // spreadsheet.element.css("width", "100%");
      // spreadsheet.resize();
      var rangeLen = this.datalen+1;

      var spreadsheet = this.$refs.spreadsheet.kendoWidget();
      var sheet = spreadsheet.activeSheet();
      var rge = "G2:G" + rangeLen;
      //alert("calc" + rge)
      sheet.range(rge).formula("C2*D2");

      var productIDRange = "A2:A" + rangeLen;
      var productNameRange= "B2:B" + rangeLen;
      var discontinuedRange = "E2:E"+ rangeLen;
      var headerCellsRange = "A1:G1";
      sheet.range(productIDRange).enable(false);
      sheet.range(productNameRange).enable(false);
      sheet.range(discontinuedRange).enable(false);
      sheet.range(headerCellsRange).enable(false);
    }
  },
  mounted: function() {
    var spreadsheet = this.$refs.spreadsheet.kendoWidget();
    spreadsheet.element.css("height", "600px");
    spreadsheet.element.css("width", "100%");
    spreadsheet.resize();

    var sheet = spreadsheet.activeSheet();
    sheet.range("F1").input("Computed via Code"); //THIS DOES NOT DISPLAY. TAKES FROM SCHEMA-MODEL
    sheet.range("G1").input("Formula");
    /*
    var rge = "G2:G78" //+ this.datalen;
    //alert(this.datalen)
     sheet.range(rge).formula("C2*D2");*/


  },

  data: function() {
    return {
      datalen: 0,
      rows: [
        {
          height: 40,
          cells: [
            {
              bold: "true",
              background: "#D5F5E3",
              textAlign: "center",
              color: "black"
            },
            {
              bold: "true",
              background: "#D5F5E3",
              textAlign: "center",
              color: "black"
            },
            {
              bold: "true",
              background: "#D5F5E3",
              textAlign: "center",
              color: "black"
            },
            {
              bold: "true",
              background: "#D5F5E3",
              textAlign: "center",
              color: "black"
            },
            {
              bold: "true",
              background: "#D5F5E3",
              textAlign: "center",
              color: "black"
            },
            {
              bold: "true",
              background: "#D5F5E3",
              textAlign: "center",
              color: "black"
            },
            {
              bold: "true",
              background: "#D5F5E3",
              textAlign: "center",
              color: "black"
            }
          ]
        }
      ],
      columns: [
        { width: 100 },
        { width: 415 },
        { width: 145 },
        { width: 145 },
        { width: 145 },
        { width: 145 },
        { width: 145 }
      ],
      datasource: {
        requestEnd: this.requestEnd,
        transport: {
          read: function(options) {
            var that = this;
            $.ajax({
              url: "https://demos.telerik.com/kendo-ui/service/Products/",
              dataType: "jsonp",
              success: function(result) {
                var res = result;
                res.forEach(element => {
                  element.Computed = element.UnitPrice * element.UnitsInStock;
                });
                options.success(res);
                that.datalen=res.length;
              },
              error: function(result) {
                options.error(result);
              }
            });
          },
          submit: function(options) {
            $.ajax({
              url: "https://demos.telerik.com/kendo-ui/service/Products/Submit",
              data: { models: kendo.stringify(e.data) },
              contentType: "application/json",
              dataType: "jsonp",
              success: function(result) {
                e.success(result.Updated, "update");
                e.success(result.Created, "create");
                e.success(result.Destroyed, "destroy");
              },
              error: function(xhr, httpStatusMessage, customErrorMessage) {
                alert(xhr.responseText);
              }
            });
          }
        },
        batch: true,
        schema: {
          model: {
            id: "ProductID",
            fields: {
              ProductID: { type: "number", editable: false},
              ProductName: { type: "string", editable: false  },
              UnitPrice: { type: "number", editable: false  },
              Discontinued: { type: "boolean", editable: false  },
              UnitsInStock: { type: "number", editable: false  },
              Computed: { type: "number", editable: false  },
              //Calculated: { type: "number" } //Does not work
            }
          }
        }
      }
    };
  },

};
</script>
<!-- <style>
	li {
		visibility: hidden
	}
</style> -->