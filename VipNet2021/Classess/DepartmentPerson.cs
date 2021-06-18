using System;
using System.Collections.Generic;
using System.Text;

namespace RegKor.Classess
{
    /// <summary>
    /// Вспомогательный класс.
    /// </summary>
    public class DepartmentPerson
    {
        private string отдел;

        /// <summary>
        /// Отдел.
        /// </summary>
        public string Отдел
        {
            get
            {
                return отдел;
            }
            set
            {
                отдел = value;
            }
        }

        private string руководитель;

        /// <summary>
        /// Руководитель.
        /// </summary>
        public string Руководитель
        {
            get
            {
                return руководитель;
            }
            set
            {
                руководитель = value;
            }
        }

    }
}
