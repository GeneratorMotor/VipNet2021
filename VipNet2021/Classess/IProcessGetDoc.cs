using System;
using System.Collections.Generic;
using System.Text;

namespace RegKor.Classess
{
    /// <summary>
    /// Способы получения документов.
    /// </summary>
    public interface IProcessGetDoc
    {
        /// <summary>
        /// Способы получения документов.
        /// </summary>
        /// <returns></returns>
        List<string> GetDoc();
    }
}
