package es.mindfm.excel;

public enum AssetCell {
     SITE_SPACE_ID(0),
        DESCRIPTION(1),
    STATUS(2),
    TYPE(3),
    ELEMENTS(4);


    private Integer position;


    public Integer getPosition()
    {
        return this.position;
    }


    private AssetCell(Integer position)
    {
        this.position = position;
    }
}
