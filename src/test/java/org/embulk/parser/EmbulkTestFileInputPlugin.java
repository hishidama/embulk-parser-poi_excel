package org.embulk.parser;

import java.nio.charset.StandardCharsets;
import java.util.List;

import org.embulk.config.ConfigDiff;
import org.embulk.config.ConfigSource;
import org.embulk.config.Task;
import org.embulk.config.TaskReport;
import org.embulk.config.TaskSource;
import org.embulk.spi.Buffer;
import org.embulk.spi.Exec;
import org.embulk.spi.FileInputPlugin;
import org.embulk.spi.TransactionalFileInput;

public class EmbulkTestFileInputPlugin implements FileInputPlugin {

	public static final String TYPE = "EmbulkTestFileInputPlugin";

	public interface PluginTask extends Task {
	}

	private List<String> list;

	public void setText(List<String> list) {
		this.list = list;
	}

	@Override
	public ConfigDiff transaction(ConfigSource config, FileInputPlugin.Control control) {
		PluginTask task = config.loadConfig(PluginTask.class);

		int taskCount = 1;
		return resume(task.dump(), taskCount, control);
	}

	@Override
	public ConfigDiff resume(TaskSource taskSource, int taskCount, FileInputPlugin.Control control) {
		control.run(taskSource, taskCount);
		return Exec.newConfigDiff();
	}

	@Override
	public void cleanup(TaskSource taskSource, int taskCount, List<TaskReport> successTaskReports) {
	}

	@Override
	public TransactionalFileInput open(TaskSource taskSource, int taskIndex) {
		return new TransactionalFileInput() {
			private boolean eof = false;
			private int index = 0;

			@Override
			public Buffer poll() {
				if (index < list.size()) {
					String s = list.get(index++) + "\n";
					return Buffer.copyOf(s.getBytes(StandardCharsets.UTF_8));
				}

				eof = true;
				return null;
			}

			@Override
			public boolean nextFile() {
				return !eof;
			}

			@Override
			public void close() {
			}

			@Override
			public void abort() {
			}

			@Override
			public TaskReport commit() {
				return Exec.newTaskReport();
			}
		};
	}
}
