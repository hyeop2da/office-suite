import { defineConfig, devices } from '@playwright/test';

// 모바일 에뮬레이션: Android Chrome(Chromium, Pixel 7), iOS Safari(WebKit, iPhone 15·iPad Pro 11, 가로 포함).
// 에뮬레이션은 화면 크기·터치·UA를 흉내 낼 뿐이며 실제 기기의 파일 선택기·다운로드 UI·가상 키보드는 시험하지 못한다.
export default defineConfig({
  testDir: 'e2e',
  timeout: 30_000,
  expect: { timeout: 7_000 },
  fullyParallel: true,
  reporter: [['list']],
  use: { baseURL: 'http://127.0.0.1:4173', trace: 'retain-on-failure', acceptDownloads: true },
  webServer: {
    command: 'npm run dev',
    url: 'http://127.0.0.1:4173/',
    reuseExistingServer: true,
    timeout: 60_000,
  },
  projects: [
    { name: 'android-chrome-pixel7', use: { ...devices['Pixel 7'] } },
    { name: 'ios-safari-iphone15', use: { ...devices['iPhone 15'] } },
    { name: 'ios-safari-iphone15-landscape', use: { ...devices['iPhone 15 landscape'] } },
    { name: 'ipados-safari-ipad-pro-11', use: { ...devices['iPad Pro 11'] } },
  ],
});
