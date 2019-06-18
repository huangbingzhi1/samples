package com.hisense.entity;

/**
 * @Author Huang.bingzhi
 * @Date 2019/6/12 22:34
 * @Version 1.0
 */
public enum Goods {
    HXDS("海信电视"),
    HXKT("海信空调"),
    KLKT("科龙空调"),
    HXBL("海信冰冷"),
    RSBL("容声冰冷"),
    XYJ("洗衣机");
    private String label;

    public String getLabel() {
        return label;
    }

    public void setLabel(String label) {
        this.label = label;
    }
    private Goods(String label) {
        this.label = label;
    }

    public static void main(String[] args) {
        Goods goods=Goods.HXBL;
        for(Goods g:Goods.values()){
            System.out.println(g.name()+":"+g.ordinal()+":"+g.label);
        }
    }
}
