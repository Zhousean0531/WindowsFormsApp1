public static class QuantityHelper
{
    public enum ProductKind
    {
        Filter,   // 濾網
        Cylinder  // 濾筒
    }

    public static string BuildQuantityText(
        ProductKind kind,
        string material,
        string qtyWeight,
        string qtyPack
    )
    {
        qtyWeight = (qtyWeight ?? "").Trim();
        qtyPack = (qtyPack ?? "").Trim();

        // 都沒填
        if (string.IsNullOrEmpty(qtyWeight) && string.IsNullOrEmpty(qtyPack))
            return "";

        // 取得單位
        var (weightUnit, packUnit) = ResolveUnit(kind, material);

        // 只填重量
        if (!string.IsNullOrEmpty(qtyWeight) && string.IsNullOrEmpty(qtyPack))
            return $"{qtyWeight} {weightUnit}";

        // 只填包數 / 板數
        if (string.IsNullOrEmpty(qtyWeight) && !string.IsNullOrEmpty(qtyPack))
            return $"{qtyPack} {packUnit}";

        // 兩個都有
        return $"{qtyWeight} {weightUnit} / {qtyPack} {packUnit}";
    }

    /// <summary>
    /// 依產品類型與料號判斷單位
    /// </summary>
    private static (string weightUnit, string packUnit)
        ResolveUnit(ProductKind kind, string material)
    {
        material = (material ?? "").Trim().ToUpperInvariant();

        if (kind == ProductKind.Filter)
        {
            // 濾網
            if (material == "IKP101")
                return ("lb", "板");

            if (material == "SI013" || material.Contains("SG"))
                return ("kg", "板");

            // 預設
            return ("kg", "板");
        }
        else
        {
            // 濾筒
            if (material == "IKP201" || material == "IKP205")
                return ("lb", "包");

            // 預設
            return ("kg", "包");
        }
    }
}
