package top.yumbo.excel.test.entity;


import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import top.yumbo.excel.annotation.business.CheckNullLogic;
import top.yumbo.excel.annotation.business.MapEntry;

import java.io.Serializable;
import java.time.LocalDateTime;

/**
 * 不动产抵押担保表
 *
 * @author yujinhua
 * @email 1015770492@qq.com
 * @date 2021-06-29 18:52:02
 */
@Data
@NoArgsConstructor
@AllArgsConstructor
public class ImportPIMCheckNullLogic implements Serializable {


    /**
     * 主键Id
     */
    private Long IMId;
    /**
     * 产品维度（字典1289）
     */
    private String pD;
    /**
     * 产品id
     */
    private Long PId;
    /**
     * 抵押人企业Id
     */
    private String bN;
    /**
     * 抵押人/所有权人名称
     */
    private String mIm;
    /**
     * 抵押标的类型(字典1181)
     * 8001301	土地使用权
     * 8001302	不动产
     * 8001303	在建工程
     */
    @MapEntry(key = "土地使用权", value = "8001301")
    @MapEntry(key = "不动产", value = "8001302")
    @MapEntry(key = "在建工程", value = "8001303")
    private String mSTIm;
    /**
     * 权证证号
     */
    private String wNIm;
    /**
     * 坐落
     */
    private String sIM;
    /**
     * 权利类型（字典1182）
     * 8001601	集体土地使用权
     * 8001602	房屋等建筑物、构筑物所有权
     * 8001603	森林、林木所有权
     * 8001604	耕地、林地、草地等土地承包经营权
     * 8001605	建设用地使用权
     * 8001606	宅基地使用权
     * 8001607	海域使用权
     * 8001608	地役权
     * 8001609	抵押权
     * 8001610	法律规定需要登记的其他不动产权利
     */
    @MapEntry(key = "集体土地使用权", value = "8001601")
    @MapEntry(key = "房屋等建筑物、构筑物所有权", value = "8001602")
    @MapEntry(key = "森林、林木所有权", value = "8001603")
    @MapEntry(key = "耕地、林地、草地等土地承包经营权", value = "8001604")
    @MapEntry(key = "建设用地使用权", value = "8001605")
    @MapEntry(key = "宅基地使用权", value = "8001606")
    @MapEntry(key = "海域使用权", value = "8001607")
    @MapEntry(key = "地役权", value = "8001608")
    @MapEntry(key = "抵押权", value = "8001609")
    @MapEntry(key = "法律规定需要登记的其他不动产权利", value = "8001610")
    private String rTIm;
    /**
     * 使用年限
     */
    private String uLIm;
    /**
     * 用途
     */
    private String pIm;
    /**
     * 是否受限(是/否，字典1022)
     */
    @MapEntry(key = "否", value = "0")
    @MapEntry(key = "是", value = "1")
    private String iLIm;
    /**
     * 房屋建筑面积,(当【抵押标的类型】是房产【8001302】或在建工程【8001303】时，必填)
     */
    @CheckNullLogic(values = {"8001302","8001303"}, follow = "mSTIm")
    private String bAIM;
    /**
     * 房屋建筑面积单位
     */
    private String bAIMUnit;
    /**
     * 宗地面积,(当【抵押标的类型】是土地使用权【8001301】时，必填)
     */
    @CheckNullLogic(follow = "mSTIm", values = "8001301")
    private String pAIM;
    /**
     * 宗地面积单位
     */
    private String pAIMUnit;
    /**
     * 是否国营主要物业或标志性资产(是/否，字典1022)
     */
    @MapEntry(key = "否", value = "0")
    @MapEntry(key = "是", value = "1")
    private String iGYYWBZZ;
    /**
     * 是否有评估价值(是/否，字典1022)
     */
    @MapEntry(key = "否", value = "0")
    @MapEntry(key = "是", value = "1")
    private String iAVIM;
    /**
     * 评估报告类型(根据字典表-项目-评估报告类型取值，【是否有评估价值】为是时，必填)
     */
    @CheckNullLogic(follow = "iAVIM", values = "1")
    private String aRTIm;
    /**
     * 评估机构(【是否有评估价值】为是时，必填)
     */
    @CheckNullLogic(follow = "iAVIM", values = "1")
    private String aMIm;
    /**
     * 评估报告名称(【是否有评估价值】为是时，必填)
     */
    @CheckNullLogic(follow = "iAVIM", values = "1")
    private String aRNIm;
    /**
     * 评估报告编号(【是否有评估价值】为是时，必填)
     */
    @CheckNullLogic(follow = "iAVIM", values = "1")
    private String aRSIm;
    /**
     * 评估价值，单位(【是否有评估价值】为是时，必填)
     */
    @CheckNullLogic(follow = "iAVIM", values = "1")
    private String aVIm;
    /**
     * 评估基准日(【是否有评估价值】为是时，必填)
     */
    @CheckNullLogic(follow = "iAVIM", values = "1")
    private String aBDIm;
    /**
     * 抵押类型(根据字典表-项目-抵押类型取值)
     */
    private String mTIM;
    /**
     * 权利性质 字典表-项目-权利性质
     */
    private String RP;
    /**
     * 担保范围（字典1310）
     */
    private String gR;
    /**
     * 担保范围描述
     */
    private String gRR;
    /**
     * 抵押权人类型（字典1311）
     */
    private String mHT;
    /**
     * 受权抵押权公司名称
     */
    private String mE;
    /**
     * 受权抵押权公司法定代表/责任人
     */
    private String mEI;
    /**
     * 受权抵押权公司联系电话
     */
    private String mETel;
    /**
     * 创建时间
     */
    private LocalDateTime gmtCreate;
    /**
     * 修改时间
     */
    private LocalDateTime gmtModified;
    /**
     * 删除标志
     */
    private String delFlag;

}
