const { spawn } = require("child_process");
const eventEmitter = require("events");

module.exports = (...args) => {
	const events = new eventEmitter();

	const command = {
		_options: {
			pipe: true
		},

		_events: events,

		_shell: undefined,

		_program: undefined,
		_argv: undefined,

		_init: (program, argv, options) => {
			options = options || {};

			Object.assign(command._options, options);

			command._program = program;
			command._argv = argv;

			return command;
		},

		start: () => {
			if(command._shell) {
				throw new Error("Command is already started");
			}

			const shell = spawn(command._program, command._argv);

			command._shell = shell;

			if(command._options.pipe) {
				shell.stdout.pipe(process.stdout);

				shell.stderr.pipe(process.stderr);

				shell.on("error", (event) => {
					events.emit("error", event);
				});

				shell.on("close", (code) => {
					command._shell = undefined;

					events.emit("close", code);
				});
			}
		},

		stop: () => {
			if(command._shell === undefined) {
				return Promise.resolve();
			}

			command._shell.kill();

			return Promise.race([
				new Promise((_resolve, reject) => {
					setTimeout(reject, 10 * 1000);
				}),
				new Promise((resolve) => {
					events.on("close", resolve);
				})
			]);
		},

		restart: () => {
			return command.stop().then(() => {
				command.start();
			});
		},

		on: (...args) => {
			events.on(...args);
		}
	};

	return command._init(...args);
};

