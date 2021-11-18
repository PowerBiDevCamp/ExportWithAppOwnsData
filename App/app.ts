import * as $ from 'jquery';

import { UiBuilder } from './services/uibuilder'
import { ViewModel, Playbook } from './services/playbook';

var resizeEmbedContainer = () => {
	var heightBuffer = 12;
	var newHeight = $(window).height() - ($("header").height() + heightBuffer);
	$("#embed-container").height(newHeight);
};

$(() => {

	var reportContainer: HTMLElement = document.getElementById("embed-container");
	var viewModel: ViewModel = window["viewModel"];

	UiBuilder.AddTopNavLink("Power BI Report", () => {
		Playbook.EmbedPowerBiReport(reportContainer, viewModel);
	});

	UiBuilder.AddTopNavLink("Paginated Report", () => {
		Playbook.EmbedPaginatedReport(reportContainer, viewModel);
	});

	UiBuilder.AddTopNavLink("Turducken", () => {
		Playbook.EmbedTurduckenReport(reportContainer, viewModel);
	});

	// intialize UI resize control
	resizeEmbedContainer();
	$(window).on("resize",resizeEmbedContainer);

});

