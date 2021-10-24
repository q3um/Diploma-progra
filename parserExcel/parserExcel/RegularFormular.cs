using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace parserExcel
{
    class RegularFormular
    {
        public readonly string CompanyNamePattern = @"(((ОБЩЕСТВО С ОГРАНИЧЕННОЙ ОТВЕТСТВЕННОСТЬЮ)|(ООО Фирма)|(ПАО)|(ООО)|(ЗАО)|(АО)|(AO)|(ОАО)|(ТОО)|(ФГУП)|(ТД НЧ)|(ТД)|(ООО ПКФ)|(МУП)|(НПК)|(МП))\s((""|«|“)\d{0,}\D{2,}\d{0,}(""|»|”),{0,2}))|^(\D{0,33}(,))|((ИП)\s\D{3,}[.])|(^[А-я]{3,14}\s[А-я]{3,14}\s[А-я]{3,14}(,)?)|^([А-я]{3,14}\s[А-я]{3,14}\s[А-я]{3,14}\s[А-я]{3,14}(,)?)|^(\D{0,}(Ltd)|(LTD))";
        public readonly string InnOrBinnPattern = @"(((ИНН)|(БИН)|(Бин)|(РНН)|(инн)|(Инн)|(ИНН:)|(БИН:)|(Бин:)|(РНН:)|(инн:)|(Инн:))/?\s{0,3}((\d{10,})|(\d{3}\s\d{3}\s\d{3}\s\d{3})))|((ИНН/КПП):?\s{0,2}\d{10}(/)\d{9})|((ИНН/КПП):?\s{0,2}\d{2,4}\s{0,2}\d{2,4}\s{0,2}\d{2,4}\s{0,2}\d{2,4}\s{0,2}/?\d{2,4}\s{0,2}\d{2,4}\s{0,2}\d{2,4})|((ИНН):?\s{0,2}\d{3}\s{0,2}\d{3}\s{0,2}\d{4})";
        public readonly string IndexPattern = @"([\s,]?\b\d{6}[,\s.])|([\s],?\d{3}-\d{3},?)|(\b\d{6}[,\s.]\b)";
        public readonly string TelephonPattern = @"((Тел)|(тел)|(моб)|(Моб)?:?[.]?\s{0,2})?((\+38|8|7|\+3|\+|\+7|\+ 7)[ ]?|\([ ]?\))?([(]?[/]?\d{3,}[/]?[)]?\s?[\- ]?)?(\d[ -]?){5,14}(\s{0,2}[(]\d{3,5}[)])?";
        public readonly string ClearKppPattern = @"(КПП):?\s{0,2}\d{9}";
        public readonly string ClearRnnPattern = @"(РНН)\s{0,2}\d{3}\s{0,2}\d{3}\s{0,2}\d{3}\s{0,2}\d{3}";
        public readonly string AcctInKpInvoice = @"\s\d{11}\s";
        public readonly string DateInKPSheet = @"\s\d{2}[.]\d{2}[.]\d{4}\s";
        public readonly string Clear = @"(/)?(,)?\s{0,2}(,)?\s{0,2}((тел)|(Тел)|(Доб)|(доб)|(Моб)|(моб)|(факс)|(факс))(:)?(.)?\s{0,2}\d{0,4}(-)?\d{0,4}(,)?";
        public readonly string Clear2 = @"($,){0,2}";

    }
}
