pipeline {
    agent {
 	   label 'windows'
}

    environment {
        RESULTS_DIR = "Results"
        SCRIPT_PATH = "AutoRunDriverScript.vbs"
        GIT_REPO_URL = 'https://github.com/lipikagupta758/FlightGUIApplication_Automation'
    }

    stages {
        stage('Checkout Code') {
            steps {
                git url: "${GIT_REPO_URL}", branch: 'main'
            }
        }

        stage('Run UFT Test') {
            steps {
                echo "Running UFT automation script..."
                bat "cscript %SCRIPT_PATH%"
            }
        }

        stage('Rename Report with Timestamp') {
            steps {
                script {
                    def timestamp = new Date().format("yyyyMMdd_HHmmss")
                    env.TIMESTAMP = timestamp
                    def reportName = "HTMLReport_${timestamp}.html"
                    env.REPORT_FILENAME = reportName
                    bat "rename ${RESULTS_DIR}\\HTMLReport.html ${reportName}"
                }
            }
        }

        stage('Publish HTML Report') {
            when {
                expression { fileExists("${RESULTS_DIR}/HTMLReport.html") }
            }
            steps {
                publishHTML(target: [
                    allowMissing: true,
                    alwaysLinkToLastBuild: true,
                    keepAll: true,
                    reportDir: "${RESULTS_DIR}",
                    reportFiles: 'HTMLReport.html',
                    reportName: 'UFT Test Report'
                ])
            }
        }

        stage('Commit Test Results to GitHub') {
            steps {
                script {
		    bat """
		    git config user.email "lipikagupta758@gmail.com"
		    git config user.name "Lipika Gupta"
                    REM Add the results directory to the Git index
                    git add Results\\HTMLReport.html
                    REM Commit changes (make sure to configure the commit message dynamically)
                    git commit -m "Add UFT test results for latest build"
                    REM Push the results to the 'Results' folder in the GitHub repository
                    git push origin main
		    """
                }
            }
        }
    }

    post {
        success {
            echo '✅ Build and test run successful!'
        }
        failure {
            echo '❌ Build or test run failed.'
        }
    }
}
