package com.hisense.entity;

/**
 * @Author Huang.bingzhi
 * @Date 2019/6/12 22:34
 * @Version 1.0
 * 商品种类
 */
public enum Goods {
    /**
     * 海信电视
     */
    HXDS("海信电视"),
    /**
     * 海信空调
     */
    HXKT("海信空调"),
    /**
     * 科龙空调
     */
    KLKT("科龙空调"),
    /**
     * 海信冰冷
     */
    HXBL("海信冰冷"),
    /**
     * 容声冰冷
     */
    RSBL("容声冰冷"),
    /**
     * 洗衣机
     */
    XYJ("洗衣机");
    private String label;

    public String getLabel() {
        return label;
    }

    public void setLabel(String label) {
        this.label = label;
    }

    Goods(String label) {
        this.label = label;
    }

    public static void main(String[] args) {
        Goods goods = Goods.HXBL;
        for (Goods g : Goods.values()) {
            System.out.println(g.name() + ":" + g.ordinal() + ":" + g.label);
        }
    }
}
