/**
 * Word's built-in equations (Insert ▸ Equation gallery) and the Equation
 * tab's structure templates, as linear format (UnicodeMath) that
 * `buildEquation` turns into OMML.
 */

export const BUILT_IN_EQUATIONS = [
  { key: "ins.eq.area", linear: "A=πr^2" },
  { key: "ins.eq.binomial", linear: "(x+a)^n=∑_(k=0)^n▒(■(n@k)) x^k a^(n−k)" },
  { key: "ins.eq.expansion", linear: "(1+x)^n=1+nx/1!+(n(n−1)x^2)/2!+⋯" },
  { key: "ins.eq.fourier", linear: "f(x)=a_0+∑_(n=1)^∞▒(a_n cos〖nπx/L〗+b_n sin〖nπx/L〗)" },
  { key: "ins.eq.pythagorean", linear: "a^2+b^2=c^2" },
  { key: "ins.eq.quadratic", linear: "x=(−b±√(b^2−4ac))/2a" },
  { key: "ins.eq.taylor", linear: "e^x=1+x/1!+x^2/2!+x^3/3!+⋯" },
  { key: "ins.eq.trig1", linear: "sin α±sin β=2 sin〖1/2 (α±β)〗 cos〖1/2 (α∓β)〗" },
  { key: "ins.eq.trig2", linear: "cos α+cos β=2 cos〖1/2 (α+β)〗 cos〖1/2 (α−β)〗" },
] as const;

export const STRUCTURES = [
  { key: "ins.eq.s.fraction", templates: ["a/b", "a∕b", "(dy)/(dx)"] },
  { key: "ins.eq.s.script", templates: ["x^2", "x_i", "x_i^2", "e^(−iωt)"] },
  { key: "ins.eq.s.radical", templates: ["√x", "√(n&x)", "∛x", "√(a^2+b^2)"] },
  { key: "ins.eq.s.integral", templates: ["∫▒f(x)dx", "∫_a^b▒f(x)dx", "∬▒f dA", "∮▒F·dr"] },
  {
    key: "ins.eq.s.largeOperator",
    templates: ["∑_(i=1)^n▒i", "∏_(i=1)^n▒i", "⋃_(i=1)^n▒A_i", "⋂_(i=1)^n▒A_i"],
  },
  { key: "ins.eq.s.bracket", templates: ["(x)", "[x]", "{x}", "|x|", "‖x‖", "⟨x⟩"] },
  { key: "ins.eq.s.function", templates: ["sin θ", "cos θ", "tan θ", "log x", "ln x"] },
  { key: "ins.eq.s.accent", templates: ["x̂", "x̃", "x⃗", "ẋ", "ẍ", "¯x"] },
  {
    key: "ins.eq.s.limit",
    templates: ["lim┬(n→∞)〖a_n〗", "log_b x", "max┬(0≤x≤1)〖f(x)〗", "min┬(0≤x≤1)〖f(x)〗"],
  },
  {
    key: "ins.eq.s.operator",
    templates: ["±", "∓", "×", "÷", "≤", "≥", "≠", "≈", "≡", "→", "⇒", "∞", "∂", "∇", "∈", "⊂"],
  },
  { key: "ins.eq.s.matrix", templates: ["[■(1&0@0&1)]", "(■(a&b@c&d))", "|■(a&b@c&d)|"] },
] as const;
