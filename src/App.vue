<template>
  <div style="width:100%;height:100%;">
        <div style="width: 100%;">
        <input style="color:Red;" type=button value="URL地址打开文档" @onclick="LoadURL()">
        </div>
    <div id="office" style="width:100%;height:80%;"></div>
  </div>
</template>
<script>
import Vue from 'vue';
// import webOfficeTpl from '../static/webOffice/iWebOffice2015.js';
import webOfficeTpl from '../public/iWebOffice2015.js';
// import { WebOffice2015 } from "../public/WebOffice.js";

export default {
  data() {
    return {
      webOffice: null,
      webOfficeObj: null
    }
  },
  beforeCreate(){
  },
  mounted(){
    console.log(webOfficeTpl);
    this.$nextTick(() => {
      this.initWebOffice();
      this.initWebOfficeObject();
    })
  },
  beforeDestroy() {
  },
  methods: {
    initWebOffice() {
      this.webOffice = new Vue({
        template: webOfficeTpl
      }).$mount('#office');
    },
    initWebOfficeObject() {
      this.webOfficeObj = new  WebOffice2015();
      this.webOfficeObj.setObj(document.getElementById('WebOffice'));
      try{
        //this.webOfficeObj.ServerUrl = "http://www.kinggrid.com:8080/iWebOffice2015/OfficeServer";
        //this.webOfficeObj.RecordID = "1551950618511";  //RecordID:本文档记录编号
        this.webOfficeObj.UserName = "XXX";
        this.webOfficeObj.FileName = "Mytemplate.doc";
        this.webOfficeObj.FileType = ".doc"; //FileType:文档类型  .doc  .xls
        this.webOfficeObj.ShowWindow = false; //显示/隐藏进度条
        this.webOfficeObj.EditType = "1"; //设置加载文档类型 0 锁定文档，1无痕迹模式，2带痕迹模式
        this.webOfficeObj.ShowMenu = 1;
        this.webOfficeObj.ShowToolBar = 0;
        this.webOfficeObj.SetCaption(this.webOfficeObj.UserName + "正在编辑文档"); // 设置控件标题栏标题文本信息
        //参数顺序依次为：控件标题栏颜色、自定义菜单开始颜色、自定义工具栏按钮开始颜色、自定义工具栏按钮结束颜色、
        //自定义工具栏按钮边框颜色、自定义工具栏开始颜色、控件标题栏文本颜色（默认值为：0x000000）
        if (!this.webOfficeObj.WebSetSkin(0xdbdbdb, 0xeaeaea, 0xeaeaea, 0xdbdbdb, 0xdbdbdb, 0xdbdbdb, 0x000000)) {
          alert(this.webOfficeObj.Status);
        }    //设置控件皮肤
        if(this.webOfficeObj.WebOpen()) {
          // StatusMsg(WebOfficeObj.Status);
        }
        this.webOfficeObj.AppendMenu("1","打开本地文件(&L)");
        this.webOfficeObj.AppendMenu("2","保存本地文件(&S)");
        this.webOfficeObj.AppendMenu("3","-");
        this.webOfficeObj.AppendMenu("4","打印预览(&C)");
        this.webOfficeObj.AppendMenu("5","退出打印预览(&E)");
        this.webOfficeObj.AddCustomMenu();
        this.webOfficeObj.HookEnabled();
        this.webOfficeObj.CreateFile() // 根据FileType设置的扩展名来创建对应的空白文档
      }
      catch(e){
        console.log("catch");
        console.log(e.description);
      }
    },
    //URL地址打开文档
    LoadURL()
    {
      console.log(this.webOfficeObj)

      try
      {

        this.webOfficeObj.ServerUrl = "http://127.0.0.1:8080/"; //服务器地址
        this.webOfficeObj.ShowMenu = 0;
        this.webOfficeObj.ShowToolBar = 0;
        SetGraySkin();			//设置控件皮肤
        if(this.webOfficeObj.WebOpen2("sample.doc"))  // 文件在服务器上的相对路径 FileName
        {
          StatusMsg(this.webOfficeObj.Status);
        }
      }
      catch(e){
        StatusMsg(e.description);
        console.log(e.description)

      }
    }


  }
}
</script>
<style lang="less">
</style>
