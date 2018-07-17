<?php

require_once __DIR__ . '/../vendor/autoload.php';

use Google\Spreadsheet\Batch\BatchRequest;
use Google\Spreadsheet\DefaultServiceRequest;
use Google\Spreadsheet\Exception\SpreadsheetNotFoundException;
use Google\Spreadsheet\ServiceRequestFactory;
use Google\Spreadsheet\SpreadsheetService;

class GoogleSpreadSheetsSender
{
    const POSITION_TO_INSERT = 1;
    protected $config;

    public function __construct()
    {
        $this->config = require __DIR__ . '/../config/config.php';
    }

    public function insert(array $rows)
    {
        $this->setServiceRequest();
        $spreadsheetService = new SpreadsheetService();
        $spreadsheetFeed = $spreadsheetService->getSpreadsheetFeed();
        try {
            $spreadsheet = $spreadsheetFeed->getById($this->config['spreadsheet_id']);
        } catch (SpreadsheetNotFoundException $exception) {
            var_dump("Spread sheet not found");
        }
        $worksheetFeed = $spreadsheet->getWorksheetFeed();

        $worksheet = $worksheetFeed->getByTitle($this->config['list']);
        $cellFeed = $worksheet->getCellFeed();
        $recordsCount = count(explode("\n", $worksheet->getCsv()));

        foreach (array_chunk($rows, 100) as $rowsChunk) {
            $batchRequest = new BatchRequest();
            $rowsChunk = array_values($rowsChunk);
            for ($rowIndex = 0; $rowIndex < count($rowsChunk); $rowIndex++) {
                $row = array_values($rowsChunk[$rowIndex]);
                for ($cell = 0; $cell <= count($row); $cell++) {
                    $batchRequest->addEntry($cellFeed->createCell(
                        ($rowIndex + $recordsCount + self::POSITION_TO_INSERT),
                        ($cell + self::POSITION_TO_INSERT),
                        $row[$cell])
                    );
                }
            }

            $cellFeed->insertBatch($batchRequest);
            usleep($this->config['insert_sleep']);
        }
    }

    protected function setServiceRequest()
    {
        $serviceRequest = new DefaultServiceRequest($this->getAccessToken());
        ServiceRequestFactory::setInstance($serviceRequest);
    }

    protected function getAccessToken()
    {
        $client = new Google_Client();
        $client->setApplicationName('Hotline parse');
        $client->setScopes([Google_Service_Sheets::SPREADSHEETS]);
        $client->setAccessType('offline');
        $client->setApprovalPrompt('force');
        $credentialsPath = __DIR__ . '/../config/credentials.json';
        if (file_exists($credentialsPath)) {
            $accessToken = json_decode(file_get_contents($credentialsPath), true);
        } else {
            // Request authorization from the user.
            $authUrl = $client->createAuthUrl();
            printf("Open the following link in your browser:\n%s\n", $authUrl);
            print 'Enter verification code: ';
            $authCode = trim(fgets(STDIN));

            // Exchange authorization code for an access token.
            $accessToken = $client->fetchAccessTokenWithAuthCode($authCode);

            // Store the credentials to disk.
            if (!file_exists(dirname($credentialsPath))) {
                mkdir(dirname($credentialsPath), 0700, true);
            }
            file_put_contents($credentialsPath, json_encode($accessToken));
            printf("Credentials saved to %s\n", $credentialsPath);
        }
        $client->setAccessToken($accessToken);

        // Refresh the token if it's expired.
        if ($client->isAccessTokenExpired()) {
            $client->fetchAccessTokenWithRefreshToken($client->getRefreshToken());
            file_put_contents($credentialsPath, json_encode($client->getAccessToken()));
        }

        return $client->getAccessToken()['access_token'];
    }
}

$sender = new GoogleSpreadSheetsSender();
$sender->insert([]);