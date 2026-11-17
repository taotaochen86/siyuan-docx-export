const D = require("../src/docx.js");
const cases = [
  "$B_1$",
  "$y=\\beta_0+\\beta_1x+\\epsilon$",
  "$\\hat{y}=\\hat{\\beta}_0+\\hat{\\beta}_1x$",
  "$E(\\epsilon)=0$",
  "$\\mathrm{Var}(\\epsilon)=\\sigma^2$",
  "$s^2=\\frac{SSE}{n-2}$",
  "$I_{CF}$",
  "$N_{2}O$",
  "$NH_{4}^{+}$",
  "$NO_{3}^{-}$",
  "$R^2$",
  "$A_1B_1$",
  "$\\cdot\\cdot\\cdot$",
];
for (const c of cases) {
  const latex = c.slice(1, -1);
  const omml = D.mathOmml(latex);
  const tags = {
    oMath: /m:oMath/.test(omml), sSub: /m:sSub/.test(omml),
    sSup: /m:sSup/.test(omml), sSubSup: /m:sSubSup/.test(omml),
    f: /<m:f>/.test(omml), acc: /m:acc/.test(omml), rad: /m:rad/.test(omml),
  };
  console.log(c + "  => " + JSON.stringify(tags));
  if (!omml) console.log("  !! 空，将回退文本");
}
