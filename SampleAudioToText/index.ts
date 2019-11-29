import { IInputs, IOutputs } from "./generated/ManifestTypes";
import * as SpeechSDK from 'microsoft-cognitiveservices-speech-sdk';

export class SampleAudioToText implements ComponentFramework.StandardControl<IInputs, IOutputs> {
	private _context: ComponentFramework.Context<IInputs>;
	private _container: HTMLDivElement;
	private _audioInput: HTMLInputElement; // for local testing only
	private _resultsArea: HTMLTextAreaElement;
	private _readAsyncBtn: HTMLButtonElement;

	//========== REAL ENVIRONMENT SCENARIO ==============
	//private _pickFileBtn: HTMLButtonElement;
	//private _audioFile: File;
	//===================================================

	private _subscriptionKey: string;
	private _subscriptionRegion: string;
	/**
	 * Empty constructor.
	 */
	constructor() {

	}

	/**
	 * Used to initialize the control instance. Controls can kick off remote server calls and other initialization actions here.
	 * Data-set values are not initialized here, use updateView.
	 * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to property names defined in the manifest, as well as utility functions.
	 * @param notifyOutputChanged A callback method to alert the framework that the control has new outputs ready to be retrieved asynchronously.
	 * @param state A piece of data that persists in one session for a single user. Can be set at any point in a controls life cycle by calling 'setControlState' in the Mode interface.
	 * @param container If a control is marked control-type='standard', it will receive an empty div element within which it can render its content.
	 */
	public init(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void, state: ComponentFramework.Dictionary, container: HTMLDivElement) {
		this._context = context;

		this.readAudioAsync = this.readAudioAsync.bind(this);

		//Provide your Subscription Key and Region
		this._subscriptionKey = "YOUR_API_KEY";
		this._subscriptionRegion = "YOUR_REGION";

		this._container = document.createElement('div');
		this._container.className = "main-container";

		container.append(this._container);

		// Add input element of file type
		// For local testing only
		this._audioInput = document.createElement('input');
		this._audioInput.type = "file";
		this._audioInput.className = "audio-input";

		this._container.appendChild(this._audioInput);

		//========== REAL ENVIRONMENT SCENARIO ==============
		// Will not work in local testing

		//this.getAudioFile = this.getAudioFile.bind(this);
		//this.convertToFile = this.convertToFile.bind(this);
		//this.setAudioFile = this.setAudioFile.bind(this);

		// this._pickFileBtn = document.createElement('button');
		// this._pickFileBtn.innerText = "Select Audio";
		// this._pickFileBtn.onclick = () => this.getAudioFile();

		//===================================================

		//Create button to trigger recognition function
		this._readAsyncBtn = document.createElement('button');
		this._readAsyncBtn.innerText = "Read Async";
		this._readAsyncBtn.onclick = () => this.readAudioAsync();

		this._container.appendChild(this._readAsyncBtn);

		let label = document.createElement('label');
		label.innerText = 'Results:';
		// Text area to show recognition results
		this._resultsArea = document.createElement("textarea");
		this._resultsArea.className = "result-area";

		this._container.appendChild(label);
		this._container.appendChild(this._resultsArea);
	}

	private readAudioAsync() {
		// for REAL ENVIRONMENT verify this._audioFile instead
		if (this._audioInput.files) {
			//disable Read Async button for the time of recognition process
			this._readAsyncBtn.disabled = true;

			// create AudioConfig from you audio file
			// only wav files supported as for now
			// local development scenario
			let audioConfig = SpeechSDK.AudioConfig.fromWavFileInput(this._audioInput.files[0]);

			//========== REAL ENVIRONMENT SCENARIO ==============
			//let audioConfig = SpeechSDK.AudioConfig.fromWavFileInput(this._audioFile);
			//===================================================

			// create speech configuration object
			// requires API key for your subscription and your subscirption region
			let speechConfig = SpeechSDK.SpeechConfig.fromSubscription(this._subscriptionKey, this._subscriptionRegion);

			// specify language for recognition
			// for list of support languages visit: https://docs.microsoft.com/en-gb/azure/cognitive-services/speech-service/language-support#speech-to-text
			speechConfig.speechRecognitionLanguage = "en-US";

			// create Speech Recognizer based on audio and speech configuration
			let speechRecognizer = new SpeechSDK.SpeechRecognizer(speechConfig, audioConfig);

			//========= OPTIONAL =============
			// this event is triggered each time when intermediate recognizing result received
			// you will receive one or more recognizing event
			// event will contain the text for the recognition since the last phrase was recognized.
			// you can comment this to just use final result in recognizeOnceAsync function
			speechRecognizer.recognizing = (recognizer, event) => {
				console.log("Recognition event:", event);
				this._resultsArea.innerHTML = event.result.text;
			};

			//function that starts recogntion in async manner
			//receives two parameters - success and error callbacks
			speechRecognizer.recognizeOnceAsync(
				(result) => {
					//success callback
					// triggers when recognition is finished and session is over
					switch (result.reason) {
						case SpeechSDK.ResultReason.RecognizedSpeech:
							// Speech recognized successfully
							this._resultsArea.innerHTML = "Result: " + result.text;
							break;
						case SpeechSDK.ResultReason.NoMatch:
							// Speech was not matched successfully. See output for more details
							let noMatchDetails = SpeechSDK.NoMatchDetails.fromResult(result);
							this._resultsArea.innerHTML = "No Match: " + SpeechSDK.NoMatchReason[noMatchDetails.reason];
							break;
						case SpeechSDK.ResultReason.Canceled:
							// Speech recognition was cancelled. See output for more details
							let cancellationDetails = SpeechSDK.CancellationDetails.fromResult(result);
							this._resultsArea.innerHTML += "Canceled: " + SpeechSDK.CancellationReason[cancellationDetails.reason];

							if (cancellationDetails.reason === SpeechSDK.CancellationReason.Error) {
								this._resultsArea.innerHTML += " Error: " + cancellationDetails.errorDetails;
							}
							break;
						default:
							break;
					}

					// Enable Read Async button after recognition is done
					this._readAsyncBtn.disabled = false;
				},
				(error) => {
					//error callback
					console.error("Error occured during recognition", error);

					// Enable Read Async button after recognition is done
					this._readAsyncBtn.disabled = false;
				}
			)

		} else {
			alert('No audio file present');
		}
	}

	//========== REAL ENVIRONMENT SCENARIO ==============
	// private getAudioFile() {
	// 	let fileOptions: ComponentFramework.DeviceApi.PickFileOptions = {
	// 		accept: "audio",
	// 		allowMultipleFiles: false,
	// 		maximumAllowedFileSize: 10000000
	// 	};

	// 	this._context.device.pickFile(fileOptions).then(
	// 		(result) => this.setAudioFile(result),
	// 		(e) => console.error("Error occured:", e)
	// 	);
	// }

	// private setAudioFile(files: ComponentFramework.FileObject[]) {
	// 	if (files) {
	// 		if (files.length > 0) {
	// 			this._audioFile = this.convertToFile(files[0]);
	// 		}
	// 	}
	// }
	//
	// private convertToFile(fileObject: ComponentFramework.FileObject): File {
	// 	return new File([fileObject.fileContent], fileObject.fileName, { type: fileObject.mimeType });
	// }
	//===================================================

	/**
	 * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
	 * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
	 */
	public updateView(context: ComponentFramework.Context<IInputs>): void {
		// Add code to update control view
	}

	/** 
	 * It is called by the framework prior to a control receiving new data. 
	 * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as “bound” or “output”
	 */
	public getOutputs(): IOutputs {
		return {};
	}

	/** 
	 * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
	 * i.e. cancelling any pending remote calls, removing listeners, etc.
	 */
	public destroy(): void {
		// Add code to cleanup control if necessary
	}
}