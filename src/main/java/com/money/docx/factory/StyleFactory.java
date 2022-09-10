package com.money.docx.factory;

import com.money.docx.constant.TbMergeEnum;
import com.money.docx.style.DocxStyle;
import org.docx4j.jaxb.Context;
import org.docx4j.wml.*;

import java.math.BigInteger;
import java.util.List;
import java.util.Optional;

/**
 * @author : money
 * @version : 1.0.0
 * @description : 风格工厂
 * @createTime : 2022-08-21 19:38:51
 */
public class StyleFactory {

    protected final static ObjectFactory factory = Context.getWmlObjectFactory();

    /**
     * 添加文本样式
     *
     * @param style 风格
     * @param r     r
     * @return {@link R}
     */
    public static R addTextStyle(R r, DocxStyle style) {
        if (style != null) {
            RPr rpr = Optional.ofNullable(r.getRPr()).orElse(factory.createRPr());
            setFontSize(rpr, style.getFontSize());
            setFontColor(rpr, style.getFontColor());
            setFontFamily(rpr, style.getFontFamily());
            if (Boolean.TRUE.equals(style.getBold())) addBoldStyle(rpr);
            if (Boolean.TRUE.equals(style.getItalic())) addUnderlineStyle(rpr);
            if (Boolean.TRUE.equals(style.getUnderline())) addItalicStyle(rpr);
            r.setRPr(rpr);
        }
        return r;
    }

    /**
     * 添加段落样式
     * 添加样式
     *
     * @param p     p
     * @param style 多克斯风格
     * @return {@link P}
     */
    public static P addParagraphStyle(P p, DocxStyle style) {
        if (style != null) {
            PPr pPr = Optional.ofNullable(p.getPPr()).orElse(factory.createPPr());
            setAlign(pPr, style.getAlign());
            p.setPPr(pPr);
        }
        return p;
    }

    /**
     * 添加单元格样式
     *
     * @param tc    tc
     * @param style 多克斯风格
     * @return {@link Tc}
     */
    public static Tc addCellStyle(Tc tc, DocxStyle style) {
        if (style != null) {
            TcPr tcPr = Optional.ofNullable(tc.getTcPr()).orElse(factory.createTcPr());
            setCellMargins(tcPr, style.getMargins());
            setCellWidth(tcPr, style.getCellWidth());
            setCellColor(tcPr, style.getCellColor());
            setCellAlign(tcPr, style.getVAlign());
            setVMerge(tcPr, style.getVMerge());
            setHMerge(tcPr, style.getHMerge());
            tc.setTcPr(tcPr);
        }
        return tc;
    }

    // ======================= style

    public static void setAlign(PPr pPr, JcEnumeration align) {
        if (align != null) {
            Jc jc = factory.createJc();
            jc.setVal(align);
            pPr.setJc(jc);
        }
    }

    public static void setCellMargins(TcPr tcPr, List<Long> margins) {
        if (margins != null && !margins.isEmpty()) {
            TcMar tcMar = new TcMar();
            // 上有下左
            for (int i = 0; i < margins.size(); i++) {
                TblWidth bW = new TblWidth();
                bW.setType("dxa");
                bW.setW(BigInteger.valueOf(margins.get(i)));
                switch (i) {
                    case 0:
                        tcMar.setTop(bW);
                    case 1:
                        tcMar.setRight(bW);
                    case 2:
                        tcMar.setBottom(bW);
                    case 3:
                        tcMar.setLeft(bW);
                }
            }
            tcPr.setTcMar(tcMar);
        }
    }

    public static void setCellColor(TcPr tcPr, String color) {
        if (color != null) {
            CTShd shd = new CTShd();
            shd.setFill(color);
            tcPr.setShd(shd);
        }
    }

    public static void setCellWidth(TcPr tcPr, Long width) {
        if (width != null) {
            TblWidth tw = factory.createTblWidth();
            tw.setType(TblWidth.TYPE_DXA);
            tw.setW(BigInteger.valueOf(width));
            tcPr.setTcW(tw);
        }
    }

    public static void setCellAlign(TcPr tcPr, STVerticalJc align) {
        if (align != null) {
            CTVerticalJc valign = new CTVerticalJc();
            valign.setVal(align);
            tcPr.setVAlign(valign);
        }
    }

    public static void setVMerge(TcPr tcPr, TbMergeEnum merge) {
        if (merge != null) {
            TcPrInner.VMerge vMerge = factory.createTcPrInnerVMerge();
            vMerge.setVal(merge.getKey());
            tcPr.setVMerge(vMerge);
        }
    }

    public static void setHMerge(TcPr tcPr, TbMergeEnum merge) {
        if (merge != null) {
            TcPrInner.HMerge hMerge = factory.createTcPrInnerHMerge();
            hMerge.setVal(merge.getKey());
            tcPr.setHMerge(hMerge);
        }
    }

    public static void setFontSize(RPr rPr, Long fontSize) {
        if (fontSize != null) {
            HpsMeasure size = new HpsMeasure();
            // 这里设置是真实word里size的一半，需要乘二
            size.setVal(BigInteger.valueOf(fontSize * 2L));
            rPr.setSz(size);
            rPr.setSzCs(size);
        }
    }

    public static void setFontColor(RPr rPr, String color) {
        if (color != null) {
            Color c = new Color();
            c.setVal(color);
            rPr.setColor(c);
        }
    }

    public static void setFontFamily(RPr rPr, String fontFamily) {
        if (fontFamily != null) {
            RFonts rf = Optional.ofNullable(rPr.getRFonts()).orElse(new RFonts());
            rPr.setRFonts(rf);
            rf.setAscii(fontFamily);
            rf.setHAnsi(fontFamily);
            rf.setAsciiTheme(null);
            rf.setHAnsiTheme(null);
        }
    }

    public static void addBoldStyle(RPr rPr) {
        BooleanDefaultTrue b = new BooleanDefaultTrue();
        b.setVal(true);
        rPr.setB(b);
    }

    public static void addItalicStyle(RPr rPr) {
        BooleanDefaultTrue b = new BooleanDefaultTrue();
        b.setVal(true);
        rPr.setI(b);
    }

    public static void addUnderlineStyle(RPr rPr) {
        U val = factory.createU();
        val.setVal(UnderlineEnumeration.SINGLE);
        rPr.setU(val);
    }
}
