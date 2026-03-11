package main

import (
	"encoding/json"
	"fmt"
	"io"
	"log"
	"net/http"
	"net/url"
	"os"
	"path/filepath"
	"strings"
)

// splitSharePointURL 解析SharePoint共享链接，提取domain, sitePath, filePath
func splitSharePointURL(sharedURL string) (domain, sitePath, filePath string, err error) {
	u, err := url.Parse(sharedURL)
	if err != nil {
		return
	}
	domain = u.Host
	parts := strings.Split(strings.Trim(u.Path, "/"), "/")
	var sitesIdx, docsIdx int = -1, -1
	for i, p := range parts {
		if p == "sites" {
			sitesIdx = i
		}
		if p == "Shared Documents" {
			docsIdx = i
		}
	}
	if sitesIdx == -1 || docsIdx == -1 {
		err = fmt.Errorf("URL格式不正确，未找到sites或Shared Documents %s", sharedURL)
		return
	}
	sitePath = strings.Join(parts[sitesIdx:sitesIdx+2], "/")
	filePath = strings.Join(parts[docsIdx+1:], "/")
	return
}

func getSiteID(domain, sitePath, token string) (string, error) {
	url := fmt.Sprintf("https://graph.microsoft.com/v1.0/sites/%s:/%s", domain, sitePath)
	req, _ := http.NewRequest("GET", url, nil)
	req.Header.Set("Authorization", "Bearer "+token)
	resp, err := http.DefaultClient.Do(req)
	if err != nil {
		return "", err
	}
	defer resp.Body.Close()
	var result map[string]interface{}

	if err := json.NewDecoder(resp.Body).Decode(&result); err != nil {
		return "", err
	}
	id, ok := result["id"].(string)
	if !ok {
		return "", fmt.Errorf("未获取到siteID: %v", result)
	}
	return id, nil
}

func getDriveID(siteID, token string) (string, error) {
	url := fmt.Sprintf("https://graph.microsoft.com/v1.0/sites/%s/drives", siteID)
	req, _ := http.NewRequest("GET", url, nil)
	req.Header.Set("Authorization", "Bearer "+token)
	resp, err := http.DefaultClient.Do(req)
	if err != nil {
		return "", err
	}
	defer resp.Body.Close()
	var result map[string]interface{}
	if err := json.NewDecoder(resp.Body).Decode(&result); err != nil {
		return "", err
	}
	value, ok := result["value"].([]interface{})
	if !ok || len(value) == 0 {
		return "", fmt.Errorf("未获取到driveID: %v", result)
	}
	drive, ok := value[0].(map[string]interface{})
	if !ok {
		return "", fmt.Errorf("未获取到driveID: %v", value[0])
	}
	id, ok := drive["id"].(string)
	if !ok {
		return "", fmt.Errorf("未获取到driveID: %v", drive)
	}
	return id, nil
}

func getFileInfo(siteID, driveID, filePath, token string) (map[string]interface{}, error) {
	url := fmt.Sprintf("https://graph.microsoft.com/v1.0/sites/%s/drives/%s/root:/%s", siteID, driveID, filePath)
	req, _ := http.NewRequest("GET", url, nil)
	req.Header.Set("Authorization", "Bearer "+token)
	resp, err := http.DefaultClient.Do(req)
	if err != nil {
		return nil, err
	}
	defer resp.Body.Close()
	var result map[string]interface{}
	if err := json.NewDecoder(resp.Body).Decode(&result); err != nil {
		return nil, err
	}
	return result, nil
}

func downloadFile(downloadURL, localPath, token string) error {
	req, _ := http.NewRequest("GET", downloadURL, nil)
	// SharePoint下载链接通常不需要token，但如果需要可加上
	// req.Header.Set("Authorization", "Bearer "+token)
	resp, err := http.DefaultClient.Do(req)
	if err != nil {
		return err
	}
	defer resp.Body.Close()
	f, err := os.Create(localPath)
	if err != nil {
		return err
	}
	defer f.Close()
	_, err = io.Copy(f, resp.Body)
	return err
}

func downloadFromSharePoint(sharedURL, saveDir, token string) error {
	domain, sitePath, filePath, err := splitSharePointURL(sharedURL)
	if err != nil {
		return err
	}
	log.Printf("domain: %s, sitePath: %s, filePath: %s", domain, sitePath, filePath)
	siteID, err := getSiteID(domain, sitePath, token)
	if err != nil {
		return err
	}
	log.Printf("siteID: %s", siteID)
	driveID, err := getDriveID(siteID, token)
	if err != nil {
		return err
	}
	log.Printf("driveID: %s", driveID)
	fileInfo, err := getFileInfo(siteID, driveID, filePath, token)
	if err != nil {
		return err
	}
	downloadURL, ok := fileInfo["@microsoft.graph.downloadUrl"].(string)
	if !ok {
		return fmt.Errorf("未获取到下载链接: %v", fileInfo)
	}
	log.Printf("downloadUrl: %s", downloadURL)
	fileName := filepath.Base(filePath)
	localPath := filepath.Join(saveDir, "downloadUrl_"+fileName)
	if err := downloadFile(downloadURL, localPath, token); err != nil {
		return err
	}
	log.Printf("下载完成: %s", localPath)
	return nil
}
