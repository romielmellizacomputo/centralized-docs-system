const cron = require("node-cron");
const command = require("../../command");

module.exports = (argv) => {
	argv = JSON.parse(JSON.stringify(argv._));

	// Remove "run command"
	argv.shift();

	// Time to run at
	const time = argv.shift();

	const program_argv = argv;
	const program = program_argv.shift();

	const valid = cron.validate(time);

	if(!valid) {
		console.error(`Cron descriptor '${time}' is not valid... `);
		process.exit(1);
	}

	cron.schedule(time, () => {
		const proc = command(program, program_argv);

		proc.on("error", (event) => {
			console.error(event);
		});

		proc.on("close", (code) => {
			if(code !== 0) {
				console.error(`Command ${program} ${program_argv.join(" ")} exited with code ${code}`);
			}
		});

		proc.start();
	});
};

