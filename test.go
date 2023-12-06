package main

import (
	"bufio"
	"fmt"
	"github.com/xuri/excelize"
	"os"
	"path/filepath"
	"sort"
	"strings"
	"sync"
	"time"
)

type StatItem struct {
	sizeInKiloBytes int64
	count           int64
}

type BySize []StatItem

func (s BySize) Len() int           { return len(s) }
func (s BySize) Swap(i, j int)      { s[i], s[j] = s[j], s[i] }
func (s BySize) Less(i, j int) bool { return s[i].sizeInKiloBytes < s[j].sizeInKiloBytes }

type ByCount []StatItem

func (c ByCount) Len() int           { return len(c) }
func (c ByCount) Swap(i, j int)      { c[i], c[j] = c[j], c[i] }
func (c ByCount) Less(i, j int) bool { return c[i].count > c[j].count }

type BySizePercentage []StatItem

func (s BySizePercentage) Len() int      { return len(s) }
func (s BySizePercentage) Swap(i, j int) { s[i], s[j] = s[j], s[i] }
func (s BySizePercentage) Less(i, j int) bool {
	totalSizeI := float64(s[i].sizeInKiloBytes * s[i].count)
	totalSizeJ := float64(s[j].sizeInKiloBytes * s[j].count)
	return totalSizeI > totalSizeJ
}

var wg sync.WaitGroup
var m map[int64]int64 = make(map[int64]int64)
var mutex sync.RWMutex

func walkDir(dir string) {
	defer wg.Done()

	err := filepath.Walk(dir, func(path string, f os.FileInfo, err error) error {
		if f.IsDir() && path != dir && err == nil {
			wg.Add(1)
			go walkDir(path)
			return filepath.SkipDir
		}

		if !f.IsDir() {
			fileSizeKB := f.Size() / 1024
			mutex.Lock()
			m[fileSizeKB] = m[fileSizeKB] + 1
			mutex.Unlock()
		}

		return err
	})

	if err != nil {
		fmt.Println("Error processing path", dir, err)
	}
}

func writeToExcel(sortedBySize, sortedByCount, sortedBySizePercentage []StatItem, filename string) {
	timestamp := time.Now().Format("2006-01-02_15-04-05")
	filename = fmt.Sprintf("FileStats_%s.xlsx", timestamp)
	f := excelize.NewFile()

	// Create Sheet Sorted by Size
	f.NewSheet("Sorted by Size")
	totalFiles, totalSize := calculateTotals(sortedBySize)
	writeStatItemsToSheet(f, "Sorted by Size", sortedBySize, totalFiles, totalSize)

	// Create Sheet Sorted by Count
	f.NewSheet("Sorted by Count")
	totalFiles, totalSize = calculateTotals(sortedByCount)
	writeStatItemsToSheet(f, "Sorted by Count", sortedByCount, totalFiles, totalSize)

	// Create Sheet Sorted by Size%
	f.NewSheet("Sorted by Size%")
	totalFiles, totalSize = calculateTotals(sortedBySizePercentage)
	writeStatItemsToSheet(f, "Sorted by Size%", sortedBySizePercentage, totalFiles, totalSize)

	// Save the file
	if err := f.SaveAs(filename); err != nil {
		fmt.Println(err)
	}
}

func calculateTotals(statItems []StatItem) (totalFiles int64, totalSize int64) {
	for _, item := range statItems {
		totalFiles += item.count
		totalSize += item.sizeInKiloBytes * item.count
	}
	return totalFiles, totalSize
}

func writeStatItemsToSheet(f *excelize.File, sheetName string, statItems []StatItem, totalFiles, totalSize int64) {
	// Названия
	f.SetCellValue(sheetName, "A1", "File Size in KB")
	f.SetCellValue(sheetName, "B1", "Count")
	f.SetCellValue(sheetName, "C1", "Size %")
	f.SetCellValue(sheetName, "D1", "Count %")

	// Общее
	f.SetCellValue(sheetName, "A2", "Total Files")
	f.SetCellValue(sheetName, "B2", totalFiles)
	f.SetCellValue(sheetName, "C2", "Total Size (KB)")
	f.SetCellValue(sheetName, "D2", totalSize)

	for i, item := range statItems {
		row := fmt.Sprintf("%d", i+3)
		f.SetCellValue(sheetName, "A"+row, item.sizeInKiloBytes)
		f.SetCellValue(sheetName, "B"+row, item.count)
		f.SetCellValue(sheetName, "C"+row, float64(item.sizeInKiloBytes*item.count)/float64(totalSize)*100)
		f.SetCellValue(sheetName, "D"+row, float64(item.count)/float64(totalFiles)*100)
	}
}

func main() {
	fmt.Print("Enter the directory path: ")
	reader := bufio.NewReader(os.Stdin)
	path, err := reader.ReadString('\n')
	if err != nil {
		fmt.Println("Error reading user input:", err)
		os.Exit(1)
	}
	path = strings.TrimSpace(path)

	wg.Add(1)
	go walkDir(path)
	wg.Wait()

	statItems := make([]StatItem, 0, len(m))
	for size, count := range m {
		statItems = append(statItems, StatItem{sizeInKiloBytes: size, count: count})
	}

	sortedBySize := make([]StatItem, len(statItems))
	copy(sortedBySize, statItems)
	sort.Sort(BySize(sortedBySize))

	sortedByCount := make([]StatItem, len(statItems))
	copy(sortedByCount, statItems)
	sort.Sort(ByCount(sortedByCount))

	sortedBySizePercentage := make([]StatItem, len(statItems))
	copy(sortedBySizePercentage, statItems)
	sort.Sort(BySizePercentage(sortedBySizePercentage))

	t := time.Now()
	formattedTime := fmt.Sprintf("%d-%02d-%02d_%02d-%02d-%02d",
		t.Year(), t.Month(), t.Day(),
		t.Hour(), t.Minute(), t.Second())
	filenameWithTimestamp := "FileStats_" + formattedTime + ".xlsx"

	writeToExcel(sortedBySize, sortedByCount, sortedBySizePercentage, filenameWithTimestamp)
}
