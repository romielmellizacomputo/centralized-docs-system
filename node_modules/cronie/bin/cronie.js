#!/usr/bin/env node

const yargs = require("yargs");

const commands = require("../src/commands");

const argv = yargs
	.command("run", "time-expression command\nRun command at specified time. ")
	.command("restart", "time-expression command\nRun command and reset it at specified time. ")
	.help()
	.alias("help", "h")
	.demandCommand(1, "You need at least one command before moving on")
	.argv;

let running = false;

process.argv.forEach(command => {
	process.argv.shift();

	if(commands[command] !== undefined && running == false) {
		running = true;
		commands[command]({
			_: JSON.parse(JSON.stringify(process.argv))
		});
	}
});

