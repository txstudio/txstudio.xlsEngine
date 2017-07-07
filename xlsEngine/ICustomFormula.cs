

namespace xlsEngine
{
    /// <summary>
    /// 實作此介面方法自訂公式計算
    /// </summary>
    public interface ICustomFormula
    {
        object GetValue(object value);
    }
}
