
<div id="lybg">
  <form action="MessageSave.Asp?MemberID=<%=mMemID%>" method="post" name="formWrite" id="formWrite">
    <div class="hengxian">
      <div class="hx-left">留言主题：</div>
      <div class="hx-right">
        <input name="MesName" type="text" id="MesName" size="40" maxlength="100" class="memberName" style="border:none; border:solid 1px #CCC; margin-left:10px"/>
        <font color="red">*</font></div>
    </div>
    <div class="hengxian">
      <div class="hx-left">留言内容：</div>
      <div class="hx-right">
        <textarea name="Content" rows="8" style="width: 380px; border:#CCC 1px solid; margin-left:10px; margin-top:5px; margin-bottom:5px" />
        </textarea>
        <font color="red">*</font></div>
    </div>
    <div class="hengxian">
      <div class="hx-left">称&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;呼：</div>
      <div class="hx-right">
        <input name="Linkman" type="text" id="Linkman" value="<%=mLinkman%>" size="20" maxlength="50" class="memberName" style="border:#CCC 1px solid; margin-left:10px"/>
        <font color="red">*</font></div>
    </div>
    <div class="hengxian">
      <div class="hx-left">性&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;别：</div>
      <div class="hx-right">
        <input name="Sex" type="radio" value="先生" <%if mSex="先生" then response.write ("checked")%> style="margin-left:10px"/>
        先生
        <input type="radio" name="Sex" value="女士" <%if mSex="女士" then response.write ("checked")%> />
        女士 <font color="red">*</font> </div>
    </div>
    <div class="hengxian" style="display:none">
      <div class="hx-left">邮政编码：</div>
      <div class="hx-right">
        <input name="ZipCode" type="text" value="000000" size="20" maxlength="20" class="memberName" />
                      <font color="red">*</font></div>
    </div>
    <div class="hengxian" style="display:none">
      <div class="hx-left">单位名称：</div>
      <div class="hx-right">
        <input name="Company" type="text" value="无" size="50" maxlength="100" class="memberName" />
                      <font color="red">*</font></div>
    </div>
    <div class="hengxian" style="display:none">
      <div class="hx-left">联系地址：</div>
      <div class="hx-right">
        <input name="Address" type="text" value="无" size="50" maxlength="100" class="memberName" />
                      <font color="red">*</font></div>
    </div>   
    <div class="hengxian">
      <div class="hx-left">联系电话：</div>
      <div class="hx-right">
        <input name="Telephone" type="text" id="Telephone" value="<%=mTelephone%>" size="20" maxlength="50" class="memberName" style="border:1px #CCC solid; margin-left:10px"/>
        <font color="red">*</font></div>
    </div>
    <div class="hengxian" style="display:none">
      <div class="hx-left">传真号码：</div>
      <div class="hx-right">
        <input name="Fax" type="text" id="Fax" value="0757" size="20" maxlength="50" class="memberName" />
                      <font color="red">*</font></div>
    </div>
    <div class="hengxian">
      <div class="hx-left">手机号码：</div>
      <div class="hx-right">
        <input name="Mobile" type="text" id="Mobile" value="<%=mMobile%>" size="20" maxlength="50" class="memberName" style="border:#CCC 1px solid; margin-left:10px"/>
        <font color="red">*</font></div>
    </div>
    <div class="hengxian">
      <div class="hx-left">电子信箱：</div>
      <div class="hx-right">
        <input name="Email" type="text" id="Email" value="<%=mEmail%>" size="30" maxlength="50" class="memberName" style="border:1px #CCC solid; margin-left:10px"/>
        <font color="red">*</font></div>
    </div>
    <div class="hengxian">
      <div class="hx-left">验&nbsp;&nbsp;证&nbsp;&nbsp;码：</div>
      <div class="hx-right">
        <input name="CheckCode" type="text" size="6" maxlength="6" class="memberName" style="border:1px #CCC solid; margin-left:10px"/>
        <a href="javascript:refreshimg()" title="看不清楚，换个图片。"><img src="Include/CheckCode/CheckCode.Asp" name="checkcode" align="absmiddle" id="checkcode" style="border: 1px solid #ffffff" /></a> <font color="red">*</font></div>
    </div>
    <div class="hengxian">
      <div class="hx-left"></div>
      <div class="hx-right"><input name="Submit" type="submit" value="发表留言" class="cartButton" style=" margin-left:10px"/></div>
    </div>
  </form>
</div>
