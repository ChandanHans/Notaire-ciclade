export const defaultConfig = {
  // API key
  apiKey: 'CAI-61BF743A2B8D2597C395BE0A60C202E7',

  // Your Developer appId, Apply in dashboard's developer section
  appId: '',

  // Is the extension enabled by default or not
  useCapsolver: true,

  // Solve captcha manually
  manualSolving: false,

  // Captcha solved callback function name
  solvedCallback: 'captchaSolvedCallback',

  // Use proxy or not
  // If useProxy is true, then proxyType, hostOrIp, port, proxyLogin, proxyPassword are required
  useProxy: false,
  proxyType: 'http',
  hostOrIp: '',
  port: '',
  proxyLogin: '',
  proxyPassword: '',

  enabledForBlacklistControl: false, // Use blacklist control
  blackUrlList: [], // Blacklist URL list

  // Is captcha enabled by default or not
  enabledForRecaptcha: true,
  enabledForRecaptchaV3: true,
  enabledForHCaptcha: true,
  enabledForFunCaptcha: true,
  enabledForImageToText: true,
  enabledForAwsCaptcha: true,

  // Task type: click or token
  reCaptchaMode: 'click',
  hCaptchaMode: 'click',

  // Delay before solving captcha
  reCaptchaDelayTime: 0,
  hCaptchaDelayTime: 0,
  textCaptchaDelayTime: 0,
  awsDelayTime: 0,

  // Number of repeated solutions after an error
  reCaptchaRepeatTimes: 10,
  reCaptcha3RepeatTimes: 10,
  hCaptchaRepeatTimes: 10,
  funCaptchaRepeatTimes: 10,
  textCaptchaRepeatTimes: 10,
  awsRepeatTimes: 10,

  // ReCaptcha V3 task type: ReCaptchaV3TaskProxyLess or ReCaptchaV3M1TaskProxyLess
  reCaptcha3TaskType: 'ReCaptchaV3TaskProxyLess',

  textCaptchaSourceAttribute: 'id="captchaImg"', // ImageToText source img's attribute name
  textCaptchaResultAttribute: 'id="CAPTCHA"', // ImageToText result element's attribute name
};
