package main

import (
	"encoding/json"
	"flag"
	"fmt"
	"io"
	"log"
	"net/http"
	"os"
	"strings"
	"time"

	"github.com/joho/godotenv"
)

type Config struct {
	TenantID       string   `json:"azure_tenant_id"`
	ClientID       string   `json:"azure_client_id"`
	ClientSecret   string   `json:"azure_client_secret"`
	SharepointURL  []string `json:"sharepoint_url"`
	ACCESS_TOKEN   string   `json:"access_token"`
	UseConfigToken bool     `json:"use_config_token"`
}

func loadConfig(path string) (*Config, error) {
	f, err := os.Open(path)
	if err != nil {
		return nil, err
	}
	defer f.Close()
	var cfg Config
	if err := json.NewDecoder(f).Decode(&cfg); err != nil {
		return nil, err
	}
	err = godotenv.Load()
	if err != nil {
		return nil, err
	}
	cfg.ClientID = os.Getenv("AZURE_CLIENT_ID")
	cfg.ClientSecret = os.Getenv("AZURE_CLIENT_SECRET")
	cfg.TenantID = os.Getenv("AZURE_TENANT_ID")
	cfg.ACCESS_TOKEN = os.Getenv("ACCESS_TOKEN")
	return &cfg, nil
}

func setupLogger() *os.File {
	logFile, err := os.OpenFile("sharepoint_downloader.log", os.O_CREATE|os.O_WRONLY|os.O_APPEND, 0644)
	if err != nil {
		log.Fatalf("无法创建日志文件: %v", err)
	}
	mw := io.MultiWriter(os.Stdout, logFile)
	log.SetOutput(mw)
	log.SetFlags(log.LstdFlags | log.Lshortfile)
	return logFile
}

func getAccessToken(cfg *Config) (string, error) {
	if cfg.UseConfigToken {
		log.Println("使用配置文件中的access_token")
		return cfg.ACCESS_TOKEN, nil
	}
	url := fmt.Sprintf("https://login.microsoftonline.com/%s/oauth2/v2.0/token", cfg.TenantID)
	data := fmt.Sprintf("client_id=%s&scope=https://graph.microsoft.com/.default&client_secret=%s&grant_type=client_credentials", cfg.ClientID, cfg.ClientSecret)
	resp, err := http.Post(url, "application/x-www-form-urlencoded", strings.NewReader(data))
	if err != nil {
		return "", err
	}
	defer resp.Body.Close()
	var result map[string]interface{}
	if err := json.NewDecoder(resp.Body).Decode(&result); err != nil {
		return "", err
	}
	if token, ok := result["access_token"].(string); ok {
		return token, nil
	}
	return "", fmt.Errorf("获取access_token失败: %v", result)
}

func main() {
	logFile := setupLogger()
	defer logFile.Close()

	// 定义命令行参数
	sharepointURL := flag.String("url", "", "SharePoint共享URL")
	saveDir := flag.String("dir", ".", "保存文件目录")
	flag.Parse()

	// 检查URL参数
	// if *sharepointURL == "" {
	// 	log.Fatalf("必须提供 -url 参数，例如: go run . -url 'https://...'")
	// }

	cfg, err := loadConfig("config.json")
	if err != nil {
		log.Fatalf("加载配置失败: %v", err)
	}
	log.Println("配置加载成功")
	downloadList := cfg.SharepointURL
	if sharepointURL != nil && *sharepointURL != "" {
		downloadList = append(downloadList, *sharepointURL)
	}
	if len(downloadList) == 0 {
		log.Fatalf("配置文件中没有sharepoint_url")
	}
	log.Printf("SharePoint URL: %v", downloadList)

	token, err := getAccessToken(cfg)
	if err != nil {
		log.Fatalf("获取access_token失败: %v", err)
	}
	log.Println("access_token获取成功")

	log.Println("access_token前30位:", token[:30])
	log.Printf("保存目录: %s\n", *saveDir)
	log.Println("开始下载")
	for _, url := range downloadList {
		err := downloadFromSharePoint(url, *saveDir, token)
		if err != nil {
			log.Printf("下载失败: %v\n", err)
		}
	}

	log.Println("程序结束", time.Now())
}
