const BASE_URL = 'https://wttr.in';

// DOM元素
const cityInput = document.getElementById('cityInput');
const searchBtn = document.getElementById('searchBtn');
const weatherCard = document.getElementById('weatherCard');
const loading = document.getElementById('loading');
const error = document.getElementById('error');
const errorMsg = document.getElementById('errorMsg');

const cityName = document.getElementById('cityName');
const weatherIcon = document.getElementById('weatherIcon');
const tempValue = document.getElementById('tempValue');
const weatherDesc = document.getElementById('weatherDesc');
const feelsLike = document.getElementById('feelsLike');
const humidity = document.getElementById('humidity');
const windSpeed = document.getElementById('windSpeed');
const visibility = document.getElementById('visibility');
const sunrise = document.getElementById('sunrise');
const sunset = document.getElementById('sunset');

// 获取天气数据
async function getWeather(city) {
  searchBtn.disabled = true;
  showLoading();
  hideError();

  try {
    const controller = new AbortController();
    const timeout = setTimeout(() => controller.abort(), 12000);

    try {
      // 请求天气数据
      const weatherRes = await fetch(
        `${BASE_URL}/${encodeURIComponent(city)}?format=j1&lang=zh`,
        { signal: controller.signal }
      );

      if (!weatherRes.ok) {
        throw new Error('获取天气数据失败，请稍后重试');
      }

      const weatherData = await weatherRes.json();

      if (!weatherData.current_condition || !weatherData.nearest_area || !weatherData.nearest_area[0]) {
        throw new Error('未找到该城市，请检查城市名称');
      }

      const area = weatherData.nearest_area[0];

      // 请求日出日落数据
      let sunData = null;
      try {
        const sunRes = await fetch(
          `https://api.sunrise-sunset.org/json?lat=${area.latitude}&lng=${area.longitude}&formatted=0`,
          { signal: controller.signal }
        );
        if (sunRes.ok) sunData = await sunRes.json();
      } catch {}

      displayWeather(weatherData, sunData, city);
    } finally {
      clearTimeout(timeout);
    }
  } catch (err) {
    if (err.name === 'AbortError') {
      showError('请求超时，请尝试其他城市或检查网络');
    } else if (err.message === 'Failed to fetch') {
      showError('网络请求失败，请检查网络连接或尝试其他城市');
    } else {
      showError(err.message);
    }
  } finally {
    hideLoading();
    searchBtn.disabled = false;
  }
}

// 渲染天气卡片
function displayWeather(weatherData, sunData, city) {
  const current = weatherData.current_condition[0];
  const area = weatherData.nearest_area[0];
  const main = getWeatherType(current.weatherDesc[0].value);
  const description = current.weatherDesc[0].value;

  weatherCard.className = 'weather-card';
  weatherCard.classList.add(main);

  cityName.textContent = city;
  weatherIcon.className = `fas ${getIconClass(description)}`;
  tempValue.textContent = current.temp_C;
  weatherDesc.textContent = description;

  feelsLike.textContent = `${current.FeelsLikeC}°C`;
  humidity.textContent = `${current.humidity}%`;
  windSpeed.textContent = `${current.windspeedKmph} km/h`;
  visibility.textContent = `${current.visibility} km`;

  // 显示日出日落，转换时区
  if (sunData && sunData.status === 'OK') {
    const tz = Math.round(parseFloat(area.longitude) / 15);
    sunrise.textContent = utcToLocal(sunData.results.sunrise, tz);
    sunset.textContent = utcToLocal(sunData.results.sunset, tz);
  } else {
    sunrise.textContent = '--:--';
    sunset.textContent = '--:--';
  }

  weatherCard.classList.remove('hidden');
}

// UTC转本地时间
function utcToLocal(isoStr, tz) {
  const d = new Date(isoStr);
  d.setHours(d.getHours() + tz);
  return `${d.getUTCHours().toString().padStart(2, '0')}:${d.getUTCMinutes().toString().padStart(2, '0')}`;
}

// 根据天气描述判断背景类型
function getWeatherType(desc) {
  const map = {
    '晴': 'clear', 'sunny': 'clear', 'clear': 'clear',
    '多云': 'clouds', 'cloudy': 'clouds', 'overcast': 'clouds',
    '雨': 'rain', 'rain': 'rain', 'drizzle': 'rain', 'thunder': 'rain',
    '雪': 'snow', 'snow': 'snow', 'blizzard': 'snow',
    '雾': 'mist', 'mist': 'mist', 'fog': 'mist', 'haze': 'mist'
  };
  for (const [key, val] of Object.entries(map)) {
    if (desc.includes(key)) return val;
  }
  return 'clouds';
}

// 根据天气描述选图标
function getIconClass(desc) {
  const map = [
    ['晴', 'sun', /sunny|clear/i],
    ['多云', 'cloud-sun', /partly cloudy|few clouds/i],
    ['阴', 'cloud', /cloudy|overcast/i],
    ['雾', 'smog', /mist|fog|haze/i],
    ['雨', 'cloud-rain', /rain|drizzle|shower/i],
    ['雷', 'bolt', /thunder|storm/i],
    ['雪', 'snowflake', /snow|blizzard|ice/i],
  ];
  for (const [, icon, regex] of map) {
    if (regex.test(desc)) return `fa-${icon}`;
  }
  return 'fa-cloud-sun';
}

function showLoading() {
  loading.classList.remove('hidden');
}

function hideLoading() {
  loading.classList.add('hidden');
}

function showError(msg) {
  errorMsg.textContent = msg;
  error.classList.remove('hidden');
}

function hideError() {
  error.classList.add('hidden');
}

// 点击搜索按钮
searchBtn.addEventListener('click', () => {
  const city = cityInput.value.trim();
  if (city) {
    getWeather(city);
  } else {
    showError('请输入城市名称');
  }
});

// 回车搜索
cityInput.addEventListener('keydown', (e) => {
  if (e.key === 'Enter') {
    searchBtn.click();
  }
});

// 快速城市标签点击
document.querySelectorAll('.city-tag').forEach(tag => {
  tag.addEventListener('click', () => {
    const city = tag.dataset.city;
    cityInput.value = city;
    getWeather(city);
  });
});
