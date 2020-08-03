package cmd

import (
	"fmt"
	"strconv"

	"github.com/360EntSecGroup-Skylar/excelize"
	"github.com/binaryfigments/nessusparse"
	"github.com/spf13/cobra"
	"github.com/spf13/viper"
)

func init() {
	rootCmd.AddCommand(checkCmd)
	checkCmd.Flags().String("filename", "", "the .nessus file to use")
	checkCmd.MarkFlagRequired("filename")
}

var checkCmd = &cobra.Command{
	Use:   "host",
	Short: "Get items per host.",
	Run: func(cmd *cobra.Command, args []string) {
		// start routine
		filename, _ := cmd.Flags().GetString("filename")
		excludeHosts := viper.GetStringSlice("exclude.hosts")
		excludePlugins := viper.GetIntSlice("exclude.plugins")

		nessus, err := nessusparse.Run(filename)
		if err != nil {
			fmt.Println(err)
		}

		f := excelize.NewFile()

		for _, host := range nessus.Report.ReportHosts {

			var hostIP string

			for _, tag := range host.HostProperties.Tags {
				// println(tag.Name)
				// println(tag.Data)
				if tag.Name == "host-ip" {
					hostIP = tag.Data
				}
			}

			// Found IP
			_, foundIP := Find(excludeHosts, hostIP)
			if foundIP {
				fmt.Println("IP address excluded: " + hostIP)
				continue
			}

			println(host.Name)
			println("host-ip: " + hostIP)

			var cell int
			cell = 2

			for index, finding := range host.ReportItems {
				// Converst pluginID string to int
				intPluginID, err := strconv.Atoi(finding.PluginID)
				if err != nil {
					fmt.Println(err)
					continue
				}
				// Look for plugins on exclude list
				_, foundPlugin := FindID(excludePlugins, intPluginID)
				if foundPlugin {
					fmt.Println("Plugin excluded: " + finding.PluginID)
					continue
				}
				if finding.RiskFactor == "None" {
					continue
				}
				// <ReportItem port="3389" svc_name="msrdp" protocol="tcp" severity="0" pluginID="10863" pluginName="SSL Certificate Information" pluginFamily="General">
				println("Port     : " + strconv.Itoa(finding.Port))
				println("Service  : " + finding.SvcName)
				println("Protocol : " + finding.Protocol)
				println(index)
				println(finding.PluginID)
				println(finding.PluginName)
				println(finding.Description)
				println(finding.Severity)
				println(finding.RiskFactor)
				println(finding.PluginOutput)

				// f := excelize.NewFile()
				f.SetCellValue("Sheet1", "A"+strconv.Itoa(cell), host.Name)
				f.SetCellValue("Sheet1", "B"+strconv.Itoa(cell), intPluginID)
				f.SetCellValue("Sheet1", "C"+strconv.Itoa(cell), finding.PluginName)
				f.SetCellValue("Sheet1", "D"+strconv.Itoa(cell), finding.Port)
				f.SetCellValue("Sheet1", "E"+strconv.Itoa(cell), finding.SvcName)
				f.SetCellValue("Sheet1", "F"+strconv.Itoa(cell), finding.Protocol)
				f.SetCellValue("Sheet1", "G"+strconv.Itoa(cell), finding.Description)
				f.SetCellValue("Sheet1", "H"+strconv.Itoa(cell), finding.Severity)
				f.SetCellValue("Sheet1", "I"+strconv.Itoa(cell), finding.RiskFactor)
				f.SetCellValue("Sheet1", "J"+strconv.Itoa(cell), finding.PluginOutput)

				// Create a new sheet.
				// index := f.NewSheet("Sheet2")
				// Set value of a cell.
				// f.SetCellValue("Sheet2", "A2", "Hello world.")
				// f.SetCellValue("Sheet1", "B2", 100)
				// Set active sheet of the workbook.
				f.SetActiveSheet(index)
				// Save xlsx file by the given path.

				cell++
			}
		}
		if err := f.SaveAs("Book1.xlsx"); err != nil {
			fmt.Println(err)
		}
		// stop
	},
}

// Find function
func Find(slice []string, val string) (int, bool) {
	for i, item := range slice {
		if item == val {
			return i, true
		}
	}
	return -1, false
}

// FindID function
func FindID(slice []int, val int) (int, bool) {
	for i, item := range slice {
		if item == val {
			return i, true
		}
	}
	return -1, false
}
