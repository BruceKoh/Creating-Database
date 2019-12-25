# ![Image description](https://github.com/mjkoh/Creating-Database/blob/master/Pic/HomeServer.png) Creating Microsoft SQL Server Express Database

Microsoft SQL Server is a database server which primarily stores and retreives data requested by other applications. These applications may run on the same or a different computer

## Installing

1. Download and install latest version of [Microsoft SQL Server Express](https://www.microsoft.com/en-us/sql-server/sql-server-editions-express)
* Current latest version is 2017 
* Choose basic version and follow the prompts
* After installation is complete, a popup will ask if you would like to install SMSS (SQL Server Management Studio). Click to install. If you have missed this step you can download [SMSS](https://docs.microsoft.com/en-us/sql/ssms/download-sql-server-management-studio-ssms?redirectedfrom=MSDN&view=sql-server-ver15) here.

2. Download and install [Anaconda](https://www.anaconda.com/distribution/)
* Choose latest Python version. Current latest version is Python 3.7. Follow the prompts and install

3. After installing Anaconda, open Anaconda Prompt and install pyodbc package

	`conda install -c anaconda pyodbc`

## Getting started
### Getting Microsoft SQL Server ready for external user access

1. Ensure that your local server is running

![Repo List](https://github.com/BruceKoh/Creating-Database/blob/master/Pic/Capture1.PNG)

2. Select Protocols for Microsoft SQL Express

![Repo List](https://github.com/BruceKoh/Creating-Database/blob/master/Pic/Capture2.PNG)

3. If TCP/IP is disabled, double click on it and set it to enabled

![Repo List](https://github.com/BruceKoh/Creating-Database/blob/master/Pic/Capture3.PNG)

4. Go to IP Addresses and ensure that all IP are active in the list

<p align="center">
  <src="https://github.com/BruceKoh/Creating-Database/blob/master/Pic/Capture4.PNG">
</p>

5. For the IPAll option, enter 1433 for TCP Port
	* Note that you can select another port number too, but SQL defaults to 1433, so we will use port 1433 here

<p align="center">
  <src="https://github.com/BruceKoh/Creating-Database/blob/master/Pic/Capture5.PNG">
</p>

6. Restart the Microsoft SQL Server to enable the changes and is ready for external user access













